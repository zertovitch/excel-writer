-- Derived from ExcelOut by Frank Schoonjans in Modula-2 - thanks!
-- First translated with Mod2Pas and P2Ada, then expanded
--
-- References to documentation are to: http://sc.openoffice.org/excelfileformat.pdf
--
-- To do:
-- =====
--  - fix Write_row_height bug (bad display on MS Excel)
--  - wrap text / text break (5.115 XF - Extended Format)
--  - border line styles (5.115 XF - Extended Format)
--  - freeze pane (5.75 PANE)
--  - BIFF > 3 and XML-based formats support
--  - ...

with Ada.Unchecked_Deallocation, Ada.Unchecked_Conversion;
with Ada.Strings.Fixed;

with Interfaces;                        use Interfaces;

-- IEEE_754 from: Simple components for Ada by Dmitry A. Kazakov
-- http://www.dmitry-kazakov.de/ada/components.htm
with IEEE_754.Long_Floats;

package body Excel_Out is

  use Ada.Streams.Stream_IO, Ada.Streams;

  -- Very low level part which deals with transferring data endian-proof,
  -- and floats in the IEEE format. This is needed for having Excel Writer
  -- totally portable on all systems and processor architectures.

  type Byte_buffer is array (Integer range <>) of Unsigned_8;

  -- Put numbers with correct endianess as bytes:
  generic
    type Number is mod <>;
    size: Positive;
  function Intel_x86_buffer( n: Number ) return Byte_buffer;
  pragma Inline(Intel_x86_buffer);

  function Intel_x86_buffer( n: Number ) return Byte_buffer is
    b: Byte_buffer(1..size);
    m: Number:= n;
  begin
    for i in b'Range loop
      b(i):= Unsigned_8(m and 255);
      m:= m / 256;
    end loop;
    return b;
  end Intel_x86_buffer;

  function Intel_16_inst is new Intel_x86_buffer( Unsigned_16, 2 );
  pragma Unreferenced (Intel_16_inst);

  function Intel_16( n: Unsigned_16 ) return Byte_buffer is
    pragma Inline(Intel_16);
  begin
    return (Unsigned_8(n and 255), Unsigned_8(Shift_Right(n, 8)));
  end Intel_16;

  function To_buf_8_bit(s: String) return Byte_buffer is
    b: Byte_buffer(s'Range);
  begin
    if s'Length > 255 then -- length doesn't fit in a byte
      raise Constraint_Error;
    end if;
    for i in b'Range loop
      b(i):= Character'Pos(s(i));
    end loop;
    return Unsigned_8(s'Length) & b;
  end To_buf_8_bit;

  function To_buf_16_bit(s: String) return Byte_buffer is
    b: Byte_buffer(s'Range);
  begin
    if s'Length > 2**16-1 then -- length doesn't fit in a U16
      raise Constraint_Error;
    end if;
    for i in b'Range loop
      b(i):= Character'Pos(s(i));
    end loop;
    return Intel_16(s'Length) & b;
  end To_buf_16_bit;

  -- Gives a byte sequence of an IEEE 64-bit number as if taken
  -- from an Intel machine (i.e. with the same endianess).
  --
  -- http://en.wikipedia.org/wiki/IEEE_754-1985#Double-precision_64_bit
  --
  function IEEE_Double_Intel_Portable(x: Long_Float) return Byte_buffer is
    pragma Inline(IEEE_Double_Intel_Portable);
    d : Byte_buffer(1..8);
    --
    use IEEE_754.Long_Floats;
    f64: constant Float_64:= To_IEEE(x);
  begin
    for i in d'Range loop
      d(i):= f64(9-i); -- Order is reversed
    end loop;
    -- Fully tested in Test_IEEE.adb
    return d;
  end IEEE_Double_Intel_Portable;

  -- Just spit the bytes of the long float - fast way.
  -- Of course this will work only on an Intel(-like) machine. We check this later.
  subtype Byte_buffer_8 is Byte_buffer(0..7);
  function IEEE_Double_Intel_Native is new
    Ada.Unchecked_Conversion(Long_Float, Byte_buffer_8);

  x: constant Long_Float:= -12345.0e-67;
  Can_use_native_IEEE: constant Boolean:=
    IEEE_Double_Intel_Portable(x) = IEEE_Double_Intel_Native(x);

  function IEEE_Double_Intel(x: Long_Float) return Byte_buffer is
    pragma Inline(IEEE_Double_Intel);
  begin
    if Can_use_native_IEEE then
      return IEEE_Double_Intel_Native(x);   -- Fast, non-portable
    else
      return IEEE_Double_Intel_Portable(x); -- Slower but portable
    end if;
  end IEEE_Double_Intel;

  -- Workaround for the severe xxx'Read xxx'Write performance
  -- problems in the GNAT and ObjectAda compilers (as in 2009)
  -- This is possible if and only if Byte = Stream_Element and
  -- arrays types are both packed and aligned the same way.
  --
  subtype Size_test_a is Byte_buffer(1..19);
  subtype Size_test_b is Ada.Streams.Stream_Element_Array(1..19);
  workaround_possible: constant Boolean:=
    Size_test_a'Size = Size_test_b'Size and
    Size_test_a'Alignment = Size_test_b'Alignment;

  procedure Block_Write(
    stream : in out Ada.Streams.Root_Stream_Type'Class;
    buffer : in     Byte_buffer
  )
  is
    pragma Inline(Block_Write);
    SE_Buffer   : Stream_Element_Array (1 .. buffer'Length);
    for SE_Buffer'Address use buffer'Address;
    pragma Import (Ada, SE_Buffer);
  begin
    if workaround_possible then
      Ada.Streams.Write(stream, SE_Buffer);
    else
      Byte_buffer'Write(stream'Access, buffer);
      -- ^ This was 30x to 70x slower on GNAT 2009
      --   Test in the Zip-Ada project.
    end if;
  end Block_Write;

  ----------------
  -- Excel BIFF --
  ----------------

  -- The original Modula-2 code counted on certain assumptions about
  -- record packing & endianess. We write data without these assumptions.

  procedure WriteBiff(
    xl     : Excel_Out_Stream'Class;
    biff_id: Unsigned_16;
    data   : Byte_buffer
  )
  is
    pragma Inline(WriteBiff);
  begin
    Block_Write(xl.xl_stream.all, Intel_16(biff_id));
    Block_Write(xl.xl_stream.all, Intel_16(Unsigned_16(data'Length)));
    Block_Write(xl.xl_stream.all, data);
  end WriteBiff;

  -- 5.8  BOF: Beginning of File
  procedure WriteBOF(xl : Excel_Out_Stream'Class) is
    data_type: constant Unsigned_16 := 16#10#;
    --  0005H = Workbook globals
    --  0006H = Visual Basic module
    --  0010H = Sheet or dialogue (see SHEETPR, S5.97)
    --  0020H = Chart
    --  0040H = Macro sheet
  begin
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#0009#, Intel_16(2) & Intel_16(data_type));
      when BIFF3 =>
        WriteBiff(xl, 16#0209#, Intel_16(2) & Intel_16(data_type) & (0,0));
    end case;
  end WriteBOF;

  -- 5.37 EOF: End of File
  procedure WriteEOF(xl : Excel_Out_Stream'Class) is
  begin
    WriteBiff(xl, 16#000A#, (1..0 => 0));
  end WriteEOF;

  -- 5.49 FORMAT (number format)
  procedure WriteFmtStr (xl : Excel_Out_Stream'Class; s : String) is
  begin
    case xl.format is
      when BIFF2 | BIFF3 =>
        WriteBiff(xl, 16#001E#, To_buf_8_bit(s));
    end case;
  end WriteFmtStr;

  -- Write built-in number formats (internal)
  procedure WriteFmtRecords (xl : Excel_Out_Stream'Class) is
    sep_1000: constant Character:= ','; -- US format
    sep_deci: constant Character:= '.'; -- US format
    -- ^ If there is any evidence of an issue with those built-in separators,
    -- we may make them configurable. NB: MS Excel 2002 and 2007 use only
    -- the index of built-in formats and discards the strings for BIFF2, but not for BIFF3...
  begin
    -- 5.12 BUILTINFMTCOUNT
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#001F#, Intel_16(Unsigned_16(last_built_in-1)));
      when BIFF3 =>
        WriteBiff(xl, 16#0056#, Intel_16(Unsigned_16(last_built_in+1)));
    end case;
    -- loop & case avoid omitting any choice
    for n in Number_format_type'First .. last_custom loop
      case n is
        when general    =>  WriteFmtStr(xl, "General");
        when decimal_0  =>  WriteFmtStr(xl, "0");
        when decimal_2  =>  WriteFmtStr(xl, "0" & sep_deci & "00"); -- 'Comma' built-in style
        when decimal_0_thousands_separator =>
          WriteFmtStr(xl, "#" & sep_1000 & "##0");
        when decimal_2_thousands_separator =>
          WriteFmtStr(xl, "#" & sep_1000 & "##0" & sep_deci & "00");
        when currency_0       =>
          WriteFmtStr(xl, "$ #" & sep_1000 & "##0;$ -#" & sep_1000 & "##0");
        when currency_red_0   =>
          WriteFmtStr(xl, "$ #" & sep_1000 & "##0;$ -#" & sep_1000 & "##0");
          -- [Red] doesn't go with non-English versions of Excel !!
        when currency_2       =>
          WriteFmtStr(xl,  "$ #" & sep_1000 & "##0" & sep_deci & "00;" &
                          "$ -#" & sep_1000 & "##0" & sep_deci & "00");
        when currency_red_2   =>
          WriteFmtStr(xl,  "$ #" & sep_1000 & "##0" & sep_deci & "00;" &
                          "$ -#" & sep_1000 & "##0" & sep_deci & "00");
        when percent_0        =>  WriteFmtStr(xl, "0%");   -- 'Percent' built-in style
        when percent_2        =>  WriteFmtStr(xl, "0" & sep_deci & "00%");
        when scientific       =>  WriteFmtStr(xl, "0" & sep_deci & "00E+00");
        when fraction_1       =>
          if xl.format >= BIFF3 then
            WriteFmtStr(xl, "#\ ?/?");
          end if;
        when fraction_2       =>
          if xl.format >= BIFF3 then
            WriteFmtStr(xl, "#\ ??/??");
          end if;
        when dd_mm_yyyy       =>  WriteFmtStr(xl, "dd/mm/yyyy");
        when dd_mmm_yy        =>  WriteFmtStr(xl, "dd/mmm/yy");
        when dd_mmm           =>  WriteFmtStr(xl, "dd/mmm");
        when mmm_yy           =>  WriteFmtStr(xl, "mmm/yy");
        when h_mm_AM_PM       =>  WriteFmtStr(xl, "h:mm\ AM/PM");
        when h_mm_ss_AM_PM    =>  WriteFmtStr(xl, "h:mm:ss\ AM/PM");
        when hh_mm            =>  WriteFmtStr(xl, "hh:mm");
        when hh_mm_ss         =>  WriteFmtStr(xl, "hh:mm:ss");
        when dd_mm_yyyy_hh_mm =>  WriteFmtStr(xl, "dd/mm/yyyy\ hh:mm");
        when percent_0_plus  =>
          WriteFmtStr(xl, "+0%;-0%;0%");
        when percent_2_plus  =>
          WriteFmtStr(xl, "+0" & sep_deci & "00%;-0" & sep_deci & "00%;0" & sep_deci & "00%");
        when date        => WriteFmtStr(xl, "yyyy\-mm\-dd");
        when date_h_m    => WriteFmtStr(xl, "yyyy\-mm\-dd\ hh:mm");
        when date_h_m_s  => WriteFmtStr(xl, "yyyy\-mm\-dd\ hh:mm:ss");
          -- !! Trouble: Excel (German Excel/French locale) writes yyyy, reads it,
          --    understands it and translates it into aaaa, but is unable to
          --    understand *our* yyyy !!
          -- Same issue as [Red] vs [Rot] above.
      end case;
    end loop;
    -- ^ Some formats in the original list caused problems, probably
    --   because of regional placeholder symbols
  end WriteFmtRecords;

  -- 5.35 DIMENSION
  procedure WriteDimensions(xl: Excel_Out_Stream'Class) is
    sheet_bounds: constant Byte_buffer:=
      Intel_16(0) &
      Intel_16(Unsigned_16(xl.maxrow)) &
      Intel_16(0) &
      Intel_16(Unsigned_16(xl.maxcolumn));
      -- 0 2 Index to first used row
      -- 2 2 Index to last used row, increased by 1
      -- 4 2 Index to first used column
      -- 6 2 Index to last used column, increased by 1
      --
      -- Since our row / column counts are 1-based, no need to increase by 1.
  begin
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#0000#, sheet_bounds);
      when BIFF3 =>
        WriteBiff(xl, 16#0200#, sheet_bounds & (0,0));
    end case;
  end WriteDimensions;

  procedure Define_number_format(
    xl           : in out Excel_Out_Stream;
    format       :    out Number_format_type;
    format_string: in     String
  )
  is
  begin
    xl.number_fmt:= xl.number_fmt + 1;
    format:= xl.number_fmt;
    WriteFmtStr(xl, format_string);
  end Define_number_format;

  procedure Write_Worksheet_header(xl : in out Excel_Out_Stream'Class) is

    procedure Define_style(fmt: Format_type; style_id: Unsigned_8) is
      Base_Level: constant:= 255;
    begin
      WriteBiff(xl,
        16#0293#,
        Intel_16(Unsigned_16(fmt) + 16#8000#) & style_id & Base_Level
      );
    end Define_style;
    --
    Comma_Style     : constant:= 3;
    Currency_Style  : constant:= 4;
    Percent_Style   : constant:= 5;
    font_for_styles : Font_type;
  begin
    WriteBOF(xl);
    -- 5.17 CODEPAGE
    WriteBiff(xl, 16#0042#, Intel_16(16#8001#)); -- Windows CP-1252
    -- 5.14 CALCMODE
    WriteBiff(xl, 16#000D#, Intel_16(1)); --  1 => automatic
    -- 5.85 REFMODE
    WriteBiff(xl, 16#000F#, Intel_16(1)); --  1 => A1 mode
    -- 5.28 DATEMODE
    WriteBiff(xl, 16#0022#, Intel_16(0)); --  0 => 1900; 1 => 1904 Date system
    -- NB: the 1904 variant is ignored by LibreOffice (<= 3.5), then wrong dates !
    --
    Define_font(xl,"Arial", 10, xl.def_font);
    Define_font(xl,"Arial", 10, font_for_styles); -- Used by BIFF3+'s styles
    WriteFmtRecords(xl);
    -- 5.111 WINDOWPROTECT
    WriteBiff(xl, 16#0019#, Intel_16(0));
    -- Define default format
    Define_format(xl, xl.def_font, general, xl.def_fmt);
    -- Define formats for the BIFF3+ "styles":
    Define_format(xl, font_for_styles, decimal_2, xl.cma_fmt);
    Define_format(xl, font_for_styles, currency_0, xl.ccy_fmt);
    Define_format(xl, font_for_styles, percent_0, xl.pct_fmt);
    -- Define styles - 5.103 STYLE p. 212
    -- NB: - it is BIFF3+ (we cheat a bit if selected format is BIFF2).
    --     - these "styles" seem to be a zombie feature of Excel 3
    --     - the whole purpose of including this is because format
    --       buttons (%)(,) in Excel 95 through 2007 are using these styles;
    --       if the styles are not defined, those buttons are not working
    --       when an Excel Writer sheet is open in MS Excel.
    Define_style(xl.cma_fmt, Comma_Style);
    Define_style(xl.ccy_fmt, Currency_Style);
    Define_style(xl.pct_fmt, Percent_Style);
    xl.dimrecpos:= Index(xl);
    WriteDimensions(xl);
    xl.is_created:= True;
  end Write_Worksheet_header;

  type Font_or_Background is (for_font, for_background);
  type Color_pair is array(Font_or_Background) of Unsigned_16;
  auto_color: constant Color_pair:=
    (16#7FFF#, -- system window text colour
     16#0019#  -- system window background colour
    );

  color_code: constant array(Excel_type, Color_type) of Color_pair :=
    ( BIFF2 =>
       (
         black      => (0, 0),
         white      => (1, 1),
         red        => (2, 2),
         green      => (3, 3),
         blue       => (4, 4),
         yellow     => (5, 5),
         magenta    => (6, 6),
         cyan       => (7, 7),
         others     => auto_color
        ),
      BIFF3 =>
        (black      => (8, 8),
         white      => (9, 9),
         red        => (10, 10),
         green      => (11, 11),
         blue       => (12, 12),
         yellow     => (13, 13),
         magenta    => (14, 14),
         cyan       => (15, 15),
         dark_red   => (16, 16),
         dark_green => (17, 17),
         dark_blue  => (18, 18),
         olive      => (19, 19),
         purple     => (20, 20),
         teal       => (21, 21),
         silver     => (22, 22),
         grey       => (23, 23),
         automatic  => auto_color
        )
     );

  -- *** Exported procedures **********************************************

  -- 5.115 XF - Extended Format
  procedure Define_format(
    xl               : in out Excel_Out_Stream;
    font             : in     Font_type;          -- Default_font(xl), or given by Define_font
    number_format    : in     Number_format_type; -- built-in, or given by Define_number_format
    cell_format      :    out Format_type;
    -- Optional parameters --
    horiz_align      : in     Horizontal_alignment:= general_alignment;
    border           : in     Cell_border:= no_border;
    shaded           : in     Boolean:= False;    -- Add a dotted background pattern
    background_color : in     Color_type:= automatic
  )
  is
    actual_number_format: Number_format_type:= number_format;
    procedure Define_BIFF2_XF is
      border_bits, mask: Unsigned_8;
    begin
      border_bits:= 0;
      mask:= 8;
      for s in Cell_border_single loop
        if border(s) then
          border_bits:= border_bits + mask;
        end if;
        mask:= mask * 2;
      end loop;
      -- 5.115.2 XF Record Contents
      WriteBiff(
        xl,
        16#0043#, -- XF code in BIFF2
        (Unsigned_8(font),
         -- ^ Index to FONT record
         0,
         -- ^ Not used
         Number_format_type'Pos(actual_number_format),
         -- ^ Number format and cell flags
         Horizontal_alignment'Pos(horiz_align) +
         border_bits +
         Boolean'Pos(shaded) * 128
         -- ^ Horizontal alignment, border style, and background
        )
      );
    end Define_BIFF2_XF;

    procedure Define_BIFF3_XF is
      area_code: Unsigned_16;
    begin
      -- 2.5.12 Patterns for Cell and Chart Background Area
      if shaded then
        area_code:=
          Boolean'Pos(shaded) * 17 +                        -- Sparse pattern, like BIFF2 "shade"
          16#40#  * color_code(BIFF3, black)(for_background) +           -- pattern colour
          16#800# * color_code(BIFF3, background_color)(for_background); -- pattern background
      elsif background_color = automatic then
        area_code:= 0;
      else
        area_code:=
          1 +                                                          -- Full pattern
          16#40#  * color_code(BIFF3, background_color)(for_background) +  -- pattern colour
          16#800# * color_code(BIFF3, background_color)(for_background);   -- pattern background
      end if;
      -- 5.115.2 XF Record Contents
      WriteBiff(
        xl,
        16#0243#, -- XF code in BIFF3
        (Unsigned_8(font),
         -- ^ 0 - Index to FONT record
         Number_format_type'Pos(actual_number_format),
         -- ^ 1 - Number format and cell flags
         0,
         -- ^ 2 - XF_TYPE_PROT (not used yet)
         16#FF#
         -- ^ 3 - XF_USED_ATTRIB
        ) &
        Intel_16(Horizontal_alignment'Pos(horiz_align)) &
        -- ^ 4 - Horizontal alignment, text break, parent style XF
        Intel_16(area_code) &
        -- ^ 6 - XF_AREA_34
        (  Boolean'Pos(border(top_single)),
           Boolean'Pos(border(left_single)),
           Boolean'Pos(border(bottom_single)),
           Boolean'Pos(border(right_single))
        )
        -- ^ 8 - XF_BORDER_34 - thin (=1) line; we could have other line styles:
        --       Thin, Medium, Dashed, Dotted, Thick, Double, Hair
      );
    end Define_BIFF3_XF;

  begin
    case xl.format is
      when BIFF2 =>
        if actual_number_format in dd_mm_yyyy .. last_custom then
          actual_number_format:= actual_number_format - 2;
        end if;
        Define_BIFF2_XF;
      when BIFF3 =>
        Define_BIFF3_XF;
    end case;
    xl.xfs:= xl.xfs + 1;
    cell_format:= Format_type(xl.xfs);
    xl.xf_def(xl.xfs):= (font => font, numb => number_format);
  end Define_format;

  procedure Header(xl : Excel_Out_Stream; page_header_string: String) is
  begin
    WriteBiff(xl, 16#0014#, To_buf_8_bit(page_header_string)); -- 5.55 p.180
  end Header;

  procedure Footer(xl : Excel_Out_Stream; page_footer_string: String) is
  begin
    WriteBiff(xl, 16#0015#, To_buf_8_bit(page_footer_string)); -- 5.48 p.173
  end Footer;

  procedure Left_Margin(xl : Excel_Out_Stream; inches: Long_Float) is
  begin
    WriteBiff(xl, 16#0026#, IEEE_Double_Intel(inches));
  end Left_Margin;

  procedure Right_Margin(xl : Excel_Out_Stream; inches: Long_Float) is
  begin
    WriteBiff(xl, 16#0027#, IEEE_Double_Intel(inches));
  end Right_Margin;

  procedure Top_Margin(xl : Excel_Out_Stream; inches: Long_Float) is
  begin
    WriteBiff(xl, 16#0028#, IEEE_Double_Intel(inches));
  end Top_Margin;

  procedure Bottom_Margin(xl : Excel_Out_Stream; inches: Long_Float) is
  begin
    WriteBiff(xl, 16#0029#, IEEE_Double_Intel(inches));
  end Bottom_Margin;

  procedure Margins(xl : Excel_Out_Stream; left, right, top, bottom: Long_Float) is
  begin
    Left_Margin(xl, left);
    Right_Margin(xl, right);
    Top_Margin(xl, top);
    Bottom_Margin(xl, bottom);
  end Margins;

  procedure Print_Row_Column_Headers(xl : Excel_Out_Stream) is
  begin
    WriteBiff(xl, 16#002A#, Intel_16(1)); -- 5.81 p.199
  end  Print_Row_Column_Headers;

  procedure Print_Gridlines(xl : Excel_Out_Stream) is
  begin
    WriteBiff(xl, 16#002B#, Intel_16(1)); -- 5.80 p.199
  end Print_Gridlines;

  procedure Page_Setup(
    xl                     : Excel_Out_Stream;
    scaling_percents       : Positive:= 100;
    fit_width_with_n_pages : Natural:= 1; -- 0: as many as possible
    fit_height_with_n_pages: Natural:= 1; -- 0: as many as possible
    orientation            : Orientation_choice:= portrait;
    scale_or_fit           : Scale_or_fit_choice:= scale
  )
  is
  begin
    -- 5.73 PAGESETUP p.192 - this is BIFF4+ (cheat if xl.format below)!
    WriteBiff(xl,
      16#00A1#,
      Intel_16(0) & -- paper type undefined
      Intel_16(Unsigned_16(scaling_percents)) &
      Intel_16(1) & -- start page number
      Intel_16(Unsigned_16(fit_width_with_n_pages)) &
      Intel_16(Unsigned_16(fit_height_with_n_pages)) &
      Intel_16(2 * Orientation_choice'Pos(orientation))
    );
    -- 5.97 SHEETPR p.207 - this is BIFF3+ (cheat if xl.format below) !
    -- NB: this field contains other informations, should be delayed
    --       in case other preferences are to be set
    WriteBiff(xl,
      16#0081#,
      Intel_16(256 * Scale_or_fit_choice'Pos(scale_or_fit))
    );
  end Page_Setup;

  y_scale: constant:= 20; -- scaling to obtain character point (pt) units

  -- 5.32 DEFAULTROWHEIGHT
  procedure Write_default_row_height (
        xl     : Excel_Out_Stream;
        height : Positive
  )
  is
    default_twips: constant Byte_buffer:= Intel_16(Unsigned_16(height * y_scale));
    options_flags: constant Byte_buffer:= (1,0);
    -- 1 = Row height and default font height do not match
  begin
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#0025#, default_twips);
      when BIFF3 =>
        WriteBiff(xl, 16#0225#, options_flags & default_twips);
    end case;
  end Write_default_row_height;

  -- 5.32 DEFCOLWIDTH
  procedure Write_default_column_width (
        xl : Excel_Out_Stream;
        width  : Positive)
  is
  begin
    WriteBiff(xl, 16#0055#, Intel_16(Unsigned_16(width)));
  end Write_default_column_width;

  procedure Write_column_width (
        xl     : Excel_Out_Stream;
        column : Positive;
        width  : Natural)
  is
  begin
    Write_column_width(xl, column, column, width);
  end Write_column_width;

  procedure Write_column_width(
    xl            : Excel_Out_Stream;
    first_column,
    last_column   : Positive;
    width         : Natural
  )
  is
  begin
    case xl.format is
      when BIFF2 =>
        -- 5.20 COLWIDTH (BIFF2 only)
        WriteBiff(xl, 16#0024#,
          Unsigned_8(first_column-1) &
          Unsigned_8(last_column-1) &
          Intel_16(Unsigned_16(width * 256)));
      when BIFF3 =>
        -- 5.18 COLINFO (BIFF3+)
        WriteBiff(xl, 16#007D#,
          Intel_16(Unsigned_16(first_column-1)) &
          Intel_16(Unsigned_16(last_column-1)) &
          Intel_16(Unsigned_16(width * 256)) &
          Intel_16(0) & -- Index to XF record (5.115) for default column formatting
          Intel_16(0) & -- Option flags
          (0,0)         -- Not used
        );
    end case;
  end Write_column_width;

  -- 5.88 ROW
  -- The OpenOffice documentation tells nice stories about row blocks,
  -- but single ROW commands can also be put before in the data stream,
  -- where the column widths are set. Excel saves with blocks of ROW
  -- commands, most of them useless.
  -- !! Bug: on Excel height /= 0 displays a blank row with no row number...

  procedure Write_row_height(
    xl : Excel_Out_Stream;
    row: Positive; height : Natural
  )
  is
    row_info_base: constant Byte_buffer:=
      Intel_16(Unsigned_16(row - 1)) &
      Intel_16(0)   & -- col. min.
      Intel_16(255) & -- col. max.
      Intel_16(Unsigned_16(height * y_scale));
  begin
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#0008#, row_info_base & (1..3 => 0) & Intel_16(0)); -- offset to data
      when BIFF3 =>
        WriteBiff(xl, 16#0208#, row_info_base & (0,0,0,0,0,1,0,0));
    end case;
  end Write_row_height;

  -- 5.45 FONT, p.171
  procedure Define_font(
    xl           : in out Excel_Out_Stream;
    font_name    :        String;
    height       :        Positive;
    font         :    out Font_type;
    style        :        Font_style:= regular;
    color        :        Color_type:= automatic
  )
  is
    style_bits, mask: Unsigned_16;
  begin
    style_bits:= 0;
    mask:= 1;
    for s in Font_style_single loop
      if style(s) then
        style_bits:= style_bits + mask;
      end if;
      mask:= mask * 2;
    end loop;
    xl.fonts:= xl.fonts + 1;
    if xl.fonts = 4 then
      xl.fonts:= 5;
      -- Anomaly! The font with index 4 is omitted in all BIFF versions.
      -- Numbering is 0, 1, 2, 3, *5*, 6,...
    end if;
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#0031#,
          Intel_16(Unsigned_16(height * y_scale)) &
          Intel_16(style_bits) &
          To_buf_8_bit(font_name)
        );
        if color /= automatic then
          -- 5.47 FONTCOLOR
          WriteBiff(xl, 16#0045#, Intel_16(color_code(BIFF2, color)(for_font)));
        end if;
      when BIFF3 =>
        WriteBiff(xl, 16#0231#,
          Intel_16(Unsigned_16(height * y_scale)) &
          Intel_16(style_bits) &
          Intel_16(color_code(BIFF3, color)(for_font)) &
          To_buf_8_bit(font_name)
        );
    end case;
    font:= Font_type(xl.fonts);
  end Define_font;

  procedure Jump_to_and_store_max(xl: in out Excel_Out_Stream; r, c: Integer) is
    pragma Inline(Jump_to_and_store_max);
  begin
    if not xl.is_created then
      raise Excel_stream_not_created;
    end if;
    Jump_to(xl, r, c); -- Store and check current position
    if r > xl.maxrow then
      xl.maxrow := r;
    end if;
    if c > xl.maxcolumn then
      xl.maxcolumn := c;
    end if;
  end Jump_to_and_store_max;

  -- 2.5.13 Cell Attributes (BIFF2 only)
  function Cell_attributes(xl: Excel_Out_Stream) return Byte_buffer is
  begin
    return
      (Unsigned_8(xl.xf_in_use),
       Unsigned_8(xl.xf_def(xl.xf_in_use).numb) + 16#40# *
       Unsigned_8(xl.xf_def(xl.xf_in_use).font),
       0
      );
  end Cell_attributes;

  function Almost_zero(x: Long_Float) return Boolean is
  begin
    return abs x <= Long_Float'Model_Small;
  end Almost_zero;

  -- Internal
  --
  -- 5.71 NUMBER
  procedure Write_as_double (
        xl     : in out Excel_Out_Stream;
        r,
        c      : Positive;
        num    : Long_Float
  )
  is
    pragma Inline(Write_as_double);
  begin
    Jump_to_and_store_max(xl, r, c);
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#0003#,
          Intel_16(Unsigned_16(r-1)) &
          Intel_16(Unsigned_16(c-1)) &
          Cell_attributes(xl) &
          IEEE_Double_Intel(num)
        );
      when BIFF3 =>
        WriteBiff(xl, 16#0203#,
          Intel_16(Unsigned_16(r-1)) &
          Intel_16(Unsigned_16(c-1)) &
          Intel_16(Unsigned_16(xl.xf_in_use)) &
          IEEE_Double_Intel(num)
        );
    end case;
    Jump_to(xl, r, c+1); -- Store and check new position
  end Write_as_double;

  -- Internal. This is BIFF2 only. BIFF format choice unchecked here.
  --
  procedure Write_as_16_bit_unsigned (
        xl : in out Excel_Out_Stream;
        r,
        c      : Positive;
        num    : Unsigned_16)
  is
    pragma Inline(Write_as_16_bit_unsigned);
  begin
    Jump_to_and_store_max(xl, r, c);
    -- 5.60 INTEGER
    WriteBiff(xl, 16#0002#,
      Intel_16(Unsigned_16(r-1)) &
      Intel_16(Unsigned_16(c-1)) &
      Cell_attributes(xl) &
      Intel_16(num)
    );
    Jump_to(xl, r, c+1); -- Store and check new position
  end Write_as_16_bit_unsigned;

  --
  -- Profile with floating-point number
  --
  procedure Write (
        xl     : in out Excel_Out_Stream;
        r,
        c      : Positive;
        num    : Long_Float
  )
  is
  begin
    case xl.format is
      when BIFF2 =>
        if num >= 0.0 and then
           num <= 65535.0 and then
           Almost_zero(num - Long_Float'Floor(num))
        then
          -- Write a 16-bit Integer (a BIFF2-only specialty),
          -- with a smaller storage :-)
          Write_as_16_bit_unsigned(xl, r, c, Unsigned_16(Long_Float'Floor(num)));
        else
          Write_as_double(xl, r, c, num);
        end if;
      when BIFF3 =>
        Write_as_double(xl, r, c, num);
    end case;
  end Write;

  --
  -- Profile with integer number
  --
  procedure Write (
        xl : in out Excel_Out_Stream;
        r,
        c      : Positive;
        num    : Integer)
  is
  begin
    case xl.format is
      when BIFF2 =>
        if num in 0..2**16-1 then
          -- We use a small storage for integers.
          -- This is a BIFF2-only specialty.
          Write_as_16_bit_unsigned(xl, r, c, Unsigned_16(num));
        else
          -- We need to us a floating-point in all other cases
          Write_as_double(xl, r, c, Long_Float(num));
        end if;
      when BIFF3 =>
        Write_as_double(xl, r, c, Long_Float(num));
    end case;
  end Write;

  -- 5.63 LABEL
  procedure Write (
        xl : in out Excel_Out_Stream;
        r,
        c      : Positive;
        str    : String)
  is
  begin
    Jump_to_and_store_max(xl, r, c);
    if str'Length > 0 then
      case xl.format is
        when BIFF2 =>
          WriteBiff(xl, 16#0004#,
            Intel_16(Unsigned_16(r-1)) &
            Intel_16(Unsigned_16(c-1)) &
            Cell_attributes(xl) &
            To_buf_8_bit(str)
          );
        when BIFF3 =>
          WriteBiff(xl, 16#0204#,
            Intel_16(Unsigned_16(r-1)) &
            Intel_16(Unsigned_16(c-1)) &
            Intel_16(Unsigned_16(xl.xf_in_use)) &
            To_buf_16_bit(str)
          );
      end case;
    end if;
    Jump_to(xl, r, c+1); -- Store and check new position
  end Write;

  procedure Write(xl: in out Excel_Out_Stream; r,c : Positive; str : Unbounded_String)
  is
  begin
    Write(xl, r,c, To_String(str));
  end Write;

  function To_Number(date: Time) return Long_Float is
  begin
    return Long_Float(date - Time_Of(1901, 01, 01, 0.0)) / 86_400.0 + 367.0;
  end To_Number;

  procedure Write(xl: in out Excel_Out_Stream; r,c : Positive; date: Time)
  is
  begin
    Write(xl, r,c, To_Number(date));
  end Write;

  -- Ada.Text_IO - like. No need to specify row & column each time
  procedure Put(xl: in out Excel_Out_Stream; num : Long_Float) is
  begin
    Write(xl, xl.curr_row, xl.curr_col, num);
  end Put;

  procedure Put(xl    : in out Excel_Out_Stream;
                num   : in Integer;
                width : in Ada.Text_IO.Field := 0; -- ignored
                base  : in Ada.Text_IO.Number_Base := 10
            )
  is
  begin
    if base = 10 then
      Write(xl, xl.curr_row, xl.curr_col, num);
    else
      declare
        use Ada.Strings.Fixed;
        s: String(1..50 + 0*width);
        -- 0*width is just to skip a warning of width being unused
        package IIO is new Ada.Text_IO.Integer_IO(Integer);
      begin
        IIO.Put(s, num, Base => base);
        Put(xl, Trim(s, Ada.Strings.Left));
      end;
    end if;
  end Put;

  procedure Put(xl: in out Excel_Out_Stream; str : String) is
  begin
    Write(xl, xl.curr_row, xl.curr_col, str);
  end Put;

  procedure Put(xl: in out Excel_Out_Stream; str : Unbounded_String) is
  begin
    Put(xl, To_String(str));
  end Put;

  procedure Put(xl: in out Excel_Out_Stream; date: Time) is
  begin
    Put(xl, To_Number(date));
  end Put;

  procedure Merge(xl: in out Excel_Out_Stream; cells : Positive) is

    -- 5.7 BLANK
    procedure Blank (
          xl : in out Excel_Out_Stream;
          r,
          c      : Positive)
    is
    begin
      Jump_to_and_store_max(xl, r, c);
      case xl.format is
        -- NB: Only with BIFF4, and only OpenOffice
        -- considers the cells really merged.
        when BIFF2 =>
          WriteBiff(xl, 16#0001#,
            Intel_16(Unsigned_16(r-1)) &
            Intel_16(Unsigned_16(c-1)) &
            Cell_attributes(xl)
          );
        when BIFF3 =>
          WriteBiff(xl, 16#0201#,
            Intel_16(Unsigned_16(r-1)) &
            Intel_16(Unsigned_16(c-1)) &
            Intel_16(Unsigned_16(xl.xf_in_use))
          );
      end case;
      Jump_to(xl, r, c+1); -- Store and check new position
    end Blank;
  begin
    for i in 1..cells loop
      Blank(xl, xl.curr_row, xl.curr_col);
    end loop;
  end Merge;

  procedure Put_Line(xl: in out Excel_Out_Stream; num : Long_Float) is
  begin
    Put(xl, num);
    New_Line(xl);
  end Put_Line;

  procedure Put_Line(xl: in out Excel_Out_Stream; num : Integer) is
  begin
    Put(xl, num);
    New_Line(xl);
  end Put_Line;

  procedure Put_Line(xl: in out Excel_Out_Stream; str : String) is
  begin
    Put(xl, str);
    New_Line(xl);
  end Put_Line;

  procedure Put_Line(xl: in out Excel_Out_Stream; str : Unbounded_String) is
  begin
    Put_Line(xl, To_String(str));
  end Put_Line;

  procedure Put_Line(xl: in out Excel_Out_Stream; date: Time) is
  begin
    Put(xl, date);
    New_Line(xl);
  end Put_Line;

  procedure New_Line(xl: in out Excel_Out_Stream; Spacing : Positive := 1) is
  begin
    Jump_to(xl, xl.curr_row + Spacing, 1);
  end New_Line;

  function Col(xl: in Excel_Out_Stream) return Positive is
  begin
    return xl.curr_col;
  end Col;

  function Column(xl: in Excel_Out_Stream) return Positive renames Col;

  function Line(xl: in Excel_Out_Stream) return Positive is
  begin
    return xl.curr_row;
  end Line;

  function Row(xl: in Excel_Out_Stream) return Positive renames Line;

  -- Relative / absolute jumps
  procedure Jump(xl: in out Excel_Out_Stream; rows, columns: Natural) is
  begin
    Jump_to(xl, xl.curr_row + rows, xl.curr_col + columns);
  end Jump;

  procedure Jump_to(xl: in out Excel_Out_Stream; row, column: Positive) is
  begin
    if row < xl.curr_row then -- trying to overwrite cells ?...
      raise Decreasing_row_index;
    end if;
    if row = xl.curr_row and then
      column < xl.curr_col
    then -- trying to overwrite cells on same row ?...
      raise Decreasing_column_index;
    end if;
    if row > 65536 then
      raise Row_out_of_range;
    elsif column > 256 then
      raise Column_out_of_range;
    end if;
    xl.curr_row:= row;
    xl.curr_col:= column;
  end Jump_to;

  procedure Next(xl: in out Excel_Out_Stream; columns: Positive:= 1) is
  begin
    Jump(xl, rows => 0, columns => columns);
  end Next;

  procedure Next_Row(xl: in out Excel_Out_Stream; rows: Positive:= 1) is
  begin
    Jump(xl, rows => rows, columns => 0);
  end Next_Row;

  procedure Use_format(
    xl           : in out Excel_Out_Stream;
    format       : in     Format_type
  )
  is
  begin
    xl.xf_in_use:= XF_Range(format);
  end Use_format;

  procedure Use_default_format(xl: in out Excel_Out_Stream) is
  begin
    Use_format(xl, xl.def_fmt);
  end Use_default_format;

  function Default_font(xl: Excel_Out_Stream) return Font_type is
  begin
    return xl.def_font;
  end Default_font;

  function Default_format(xl: Excel_Out_Stream) return Format_type is
  begin
    return xl.def_fmt;
  end Default_format;

  procedure Reset(
    xl           : in out Excel_Out_Stream'Class;
    excel_format :        Excel_type:= Default_Excel_type
  )
  is
    dummy_xl_with_defaults: Excel_Out_Pre_Root_Type;
  begin
    -- Check if we are trying to re-use a half-finished object (ouch!):
    if xl.is_created and not xl.is_closed then
      raise Excel_stream_not_closed;
    end if;
    -- We will reset evything with defaults, except this:
    dummy_xl_with_defaults.format:= excel_format;
    -- Now we reset xl:
    Excel_Out_Pre_Root_Type(xl):= dummy_xl_with_defaults;
  end Reset;

  procedure Finish(xl : in out Excel_Out_Stream'Class) is
  begin
    WriteEOF(xl);
    Set_Index(xl, xl.dimrecpos); -- Go back to overwrite the DIMENSION record with correct data
    WriteDimensions(xl);
    xl.is_closed:= True;
  end Finish;

  ----------------------
  -- Output to a file --
  ----------------------

  procedure Create(
    xl           : in out Excel_Out_File;
    file_name    :        String;
    excel_format :        Excel_type:= Default_Excel_type
  )
  is
  begin
    Reset(xl, excel_format);
    xl.xl_file:= new Ada.Streams.Stream_IO.File_Type;
    Create(xl.xl_file.all, Out_File, file_name);
    xl.xl_stream:= XL_Raw_Stream_Class(Stream(xl.xl_file.all));
    Write_Worksheet_header(xl);
  end Create;

  procedure Close(xl : in out Excel_Out_File) is
    procedure Dispose is new
      Ada.Unchecked_Deallocation(Ada.Streams.Stream_IO.File_Type, XL_file_acc);
  begin
    Finish(xl);
    Close(xl.xl_file.all);
    Dispose(xl.xl_file);
  end Close;

  -- Set the index on the file
  procedure Set_Index (xl: in out Excel_Out_File;
                       To: Ada.Streams.Stream_IO.Positive_Count)
  is
  begin
    Ada.Streams.Stream_IO.Set_Index(xl.xl_file.all, To);
  end Set_Index;

  -- Return the index of the file
  function Index (xl: Excel_Out_File) return Ada.Streams.Stream_IO.Count
  is
  begin
    return Ada.Streams.Stream_IO.Index(xl.xl_file.all);
  end Index;

  function Is_Open(xl : in Excel_Out_File) return Boolean is
  begin
    if xl.xl_file = null then
      return False;
    end if;
    return Ada.Streams.Stream_IO.Is_Open(xl.xl_file.all);
  end Is_Open;

  ------------------------
  -- Output to a string --
  ------------------------
  -- Code reused from Zip_Streams

  procedure Read
    (Stream : in out Unbounded_Stream;
     Item   : out Stream_Element_Array;
     Last   : out Stream_Element_Offset) is
  begin
    -- Item is read from the stream. If (and only if) the stream is
    -- exhausted, Last will be < Item'Last. In that case, T'Read will
    -- raise an End_Error exception.
    --
    -- Cf: RM 13.13.1(8), RM 13.13.1(11), RM 13.13.2(37) and
    -- explanations by Tucker Taft
    --
    Last:= Item'First - 1;
    -- if Item is empty, the following loop is skipped; if Stream.Loc
    -- is already indexing out of Stream.Unb, that value is also appropriate
    for i in Item'Range loop
      Item(i) := Character'Pos (Element(Stream.Unb, Stream.Loc));
      Stream.Loc := Stream.Loc + 1;
      Last := i;
    end loop;
  exception
    when Ada.Strings.Index_Error =>
      null; -- what could be read has been read; T'Read will raise End_Error
  end Read;

  procedure Write
    (Stream : in out Unbounded_Stream;
     Item   : Stream_Element_Array) is
  begin
    for I in Item'Range loop
      if Length(Stream.Unb) < Stream.Loc then
        Append(Stream.Unb, Character'Val(Item(I)));
      else
        Replace_Element(Stream.Unb, Stream.Loc, Character'Val(Item(I)));
      end if;
      Stream.Loc := Stream.Loc + 1;
    end loop;
  end Write;

  procedure Set_Index (S : access Unbounded_Stream; To : Positive) is
  begin
    if Length(S.Unb) < To then
      for I in Length(S.Unb) .. To loop
        Append(S.Unb, ASCII.NUL);
      end loop;
    end if;
    S.Loc := To;
  end Set_Index;

  function Index (S : access Unbounded_Stream) return Integer is
  begin
    return S.Loc;
  end Index;

  --- ***

  procedure Create(
    xl           : in out Excel_Out_String;
    excel_format :        Excel_type:= Default_Excel_type
  )
  is
  begin
    Reset(xl, excel_format);
    xl.xl_memory:= new Unbounded_Stream;
    xl.xl_memory.Unb:= Null_Unbounded_String;
    xl.xl_memory.Loc:= 1;
    xl.xl_stream:= XL_Raw_Stream_Class(xl.xl_memory);
    Write_Worksheet_header(xl);
  end Create;

  procedure Close(xl : in out Excel_Out_String) is
  begin
    Finish(xl);
  end Close;

  function Contents(xl: Excel_Out_String) return String is
  begin
    if not xl.is_closed then
      raise Excel_stream_not_closed;
    end if;
    return To_String(xl.xl_memory.Unb);
  end Contents;

  -- Set the index on the Excel string stream
  procedure Set_Index (xl: in out Excel_Out_String;
                       To: Ada.Streams.Stream_IO.Positive_Count)
  is
  begin
    Set_Index(xl.xl_memory, Integer(To));
  end Set_Index;

  -- Return the index of the Excel string stream
  function Index (xl: Excel_Out_String) return Ada.Streams.Stream_IO.Count
  is
  begin
    return Ada.Streams.Stream_IO.Count(Index(xl.xl_memory));
  end Index;

  function "&"(a,b: Font_style) return Font_style is
  begin
    return a or b; -- "or" is predefined for sets (=array of Boolean)
  end "&";

  function "&"(a,b: Cell_border) return Cell_border is
  begin
    return a or b; -- "or" is predefined for sets (=array of Boolean)
  end "&";

end Excel_Out;
