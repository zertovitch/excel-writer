--  References to documentation are to: http://www.openoffice.org/sc/excelfileformat.pdf
--
--  To do:
--  =====
--  - Complete Write_as_Formula
--  - Unicode (for binary Excel: requires BIFF8, but BIFF8 is pretty difficult)
--  - border line styles (5.115 XF - Extended Format)
--  - XML-based formats support (XLSX, ODS, ...)
--  - ...

with Ada.Unchecked_Conversion,
     Ada.Unchecked_Deallocation,
     Ada.Strings.Fixed;

with Interfaces;

with Excel_Out.IEEE_754.Generic_Double_Precision;

package body Excel_Out is

  use Ada.Strings.Unbounded, Interfaces;

  --  Very low level part which deals with transferring data in an endian-neutral way,
  --  and floats in the IEEE format. This is needed for having Excel Writer
  --  totally portable on all systems and processor architectures.

  type Byte_Buffer is array (Integer range <>) of Unsigned_8;
  empty_buffer : constant Byte_Buffer := (1 .. 0 => 0);

  --  Put numbers with correct endianess as bytes:
  generic
    type Number is mod <>;
    size : Positive;
  function Intel_x86_buffer (n : Number) return Byte_Buffer;
  pragma Inline (Intel_x86_buffer);

  function Intel_x86_buffer (n : Number) return Byte_Buffer is
    b : Byte_Buffer (1 .. size);
    m : Number := n;
  begin
    for i in b'Range loop
      b (i) := Unsigned_8 (m and 255);
      m := m / 256;
    end loop;
    return b;
  end Intel_x86_buffer;

  function Intel_32 is new Intel_x86_buffer (Unsigned_32, 4);

  function Intel_16 (n : Unsigned_16) return Byte_Buffer is
    pragma Inline (Intel_16);
  begin
    return (Unsigned_8 (n and 255), Unsigned_8 (Shift_Right (n, 8)));
  end Intel_16;

  --  2.5.2 Byte Strings, 8-bit string length (BIFF2-BIFF5), p. 187
  function To_buf_8_bit_length (s : String) return Byte_Buffer is
    b : Byte_Buffer (s'Range);
  begin
    if s'Length > 255 then -- length doesn't fit in a byte
      raise Constraint_Error;
    end if;
    for i in b'Range loop
      b (i) := Character'Pos (s (i));
    end loop;
    return Unsigned_8 (s'Length) & b;
  end To_buf_8_bit_length;

  --  2.5.2 Byte Strings, 16-bit string length (BIFF2-BIFF5), p. 187
  function To_buf_16_bit_length (s : String) return Byte_Buffer is
    b : Byte_Buffer (s'Range);
  begin
    if s'Length > 2**16 - 1 then -- length doesn't fit in a 16-bit number
      raise Constraint_Error;
    end if;
    for i in b'Range loop
      b (i) := Character'Pos (s (i));
    end loop;
    return Intel_16 (s'Length) & b;
  end To_buf_16_bit_length;

  --  --  2.5.3 Unicode Strings, 16-bit string length (BIFF2-BIFF5), p. 17
  --  function To_buf_16_bit_length(s: Wide_String) return Byte_buffer is
  --    b: Byte_buffer(1 .. 2 * s'Length);
  --    j: Integer:= 1;
  --  begin
  --    if s'Length > 2**16-1 then -- length doesn't fit in a 16-bit number
  --      raise Constraint_Error;
  --    end if;
  --    for i in s'Range loop
  --      b(j)  := Unsigned_8(Unsigned_32'(Wide_Character'Pos(s(i))) and 255);
  --      b(j+1):= Unsigned_8(Shift_Right(Unsigned_32'(Wide_Character'Pos(s(i))), 8));
  --      j:= j + 2;
  --    end loop;
  --    return
  --      Intel_16(s'Length) &
  --      (1 => 1) &  --  Character compression (ccompr): 1 = Uncompressed (16-bit characters)
  --      b;
  --  end To_buf_16_bit_length;

  --  Gives a byte sequence of an IEEE 64-bit number as if taken
  --  from an Intel machine (i.e. with the same endianess).
  --
  --  http://en.wikipedia.org/wiki/IEEE_754-1985#Double-precision_64_bit
  --

   package IEEE_LF is new IEEE_754.Generic_Double_Precision (Long_Float);

  function IEEE_Double_Intel_Portable (x : Long_Float) return Byte_Buffer is
    pragma Inline (IEEE_Double_Intel_Portable);
    d : Byte_Buffer (1 .. 8);
    --
    f64 : constant IEEE_LF.Float_64 := IEEE_LF.To_IEEE (x);
  begin
    for i in d'Range loop
      d (i) := f64 (9 - i); -- Order is reversed
    end loop;
    --  Fully tested in Test_IEEE.adb
    return d;
  end IEEE_Double_Intel_Portable;

  --  Just spit the bytes of the long float - fast way.
  --  Of course this will work only on an Intel(-like) machine. We check this later.
  subtype Byte_buffer_8 is Byte_Buffer (0 .. 7);
  function IEEE_Double_Intel_Native is new
    Ada.Unchecked_Conversion (Long_Float, Byte_buffer_8);

  x_test : constant Long_Float := -12345.0e-67;
  Can_use_native_IEEE : constant Boolean :=
    IEEE_Double_Intel_Portable (x_test) = IEEE_Double_Intel_Native (x_test);

  function IEEE_Double_Intel (x : Long_Float) return Byte_Buffer is
    pragma Inline (IEEE_Double_Intel);
  begin
    if Can_use_native_IEEE then
      return IEEE_Double_Intel_Native (x);   -- Fast, non-portable
    else
      return IEEE_Double_Intel_Portable (x); -- Slower but portable
    end if;
  end IEEE_Double_Intel;

  --  Workaround for the severe xxx'Read xxx'Write performance
  --  problems in the GNAT and ObjectAda compilers (as in 2009)
  --  This is possible if and only if Byte = Stream_Element and
  --  arrays types are both packed and aligned the same way.
  --
  subtype Size_test_a is Byte_Buffer (1 .. 19);
  subtype Size_test_b is Ada.Streams.Stream_Element_Array (1 .. 19);
  workaround_possible : constant Boolean :=
    Size_test_a'Size = Size_test_b'Size and
    Size_test_a'Alignment = Size_test_b'Alignment;

  procedure Block_Write (
    stream : in out Ada.Streams.Root_Stream_Type'Class;
    buffer : in     Byte_Buffer
  )
  is
    pragma Inline (Block_Write);
    SE_Buffer   : Ada.Streams.Stream_Element_Array (1 .. buffer'Length);
    for SE_Buffer'Address use buffer'Address;
    pragma Import (Ada, SE_Buffer);
  begin
    if workaround_possible then
      Ada.Streams.Write (stream, SE_Buffer);
    else
      Byte_Buffer'Write (stream'Access, buffer);
      --  ^ This was 30x to 70x slower on GNAT 2009
      --    Test in the Zip-Ada project.
    end if;
  end Block_Write;

  ----------------
  -- Excel BIFF --
  ----------------

  --  The original Modula-2 code counted on certain assumptions about
  --  record packing & endianess. We write data without these assumptions.

  procedure Write_Biff (
    xl     : Excel_Out_Stream'Class;
    biff_id : Unsigned_16;
    data   : Byte_Buffer
  )
  is
    pragma Inline (Write_Biff);
  begin
    Block_Write (xl.xl_stream.all, Intel_16 (biff_id));
    Block_Write (xl.xl_stream.all, Intel_16 (Unsigned_16 (data'Length)));
    Block_Write (xl.xl_stream.all, data);
  end Write_Biff;

  --  5.8  BOF: Beginning of File, p.135
  procedure Write_BOF (xl : Excel_Out_Stream'Class) is

    function BOF_suffix return Byte_Buffer is  --  5.8.1 Record BOF
    begin
      case xl.xl_format is
        when BIFF2 =>
          return empty_buffer;
        when BIFF3 | BIFF4 =>
          return (0, 0);  --  Not used
        --  when BIFF8 =>
        --    return (1,1,1,1);
      end case;
    end BOF_suffix;

    --  0005H = Workbook globals
    --  0006H = Visual Basic module
    --  0010H = Sheet or dialogue (see SHEETPR, S5.97)
    Sheet_or_dialogue : constant := 16#10#;
    --  0020H = Chart
    --  0040H = Macro sheet
    biff_record_identifier : constant array (Excel_Type) of Unsigned_16 :=
      (BIFF2 => 16#0009#,
       BIFF3 => 16#0209#,
       BIFF4 => 16#0409#
       --  BIFF8 => 16#0809#
      );
    biff_version : constant array (Excel_Type) of Unsigned_16 :=
      (BIFF2 => 16#0200#,
       BIFF3 => 16#0300#,
       BIFF4 => 16#0400#
       --  BIFF8 => 16#0600#
      );
  begin
    Write_Biff (xl,
      biff_record_identifier (xl.xl_format),
      Intel_16 (biff_version (xl.xl_format)) &
      Intel_16 (Sheet_or_dialogue) &
      BOF_suffix
    );
  end Write_BOF;

  --  5.49 FORMAT (actually, a number format; the full format is called XF (extended format))
  procedure Write_Number_Format_String (xl : Excel_Out_Stream'Class; s : String) is
  begin
    case xl.xl_format is
      when BIFF2 | BIFF3 =>
        Write_Biff (xl, 16#001E#, To_buf_8_bit_length (s));
      when BIFF4 =>
        Write_Biff (xl, 16#041E#, (0, 0) & To_buf_8_bit_length (s));
     --   when BIFF8 =>
     --     Write_Biff(xl, 16#041E#, (0, 0) &  --  should be: format index used in other records
     --       To_buf_8_bit_length(s));
    end case;
  end Write_Number_Format_String;

  --  Write built-in number formats (internal)
  procedure Write_Default_Number_Formats (xl : Excel_Out_Stream'Class) is
    sep_1000 : constant Character := ','; -- US format
    sep_deci : constant Character := '.'; -- US format
    --  ^ If there is any evidence of an issue with those built-in separators,
    --  we may make them configurable. NB: MS Excel 2002 and 2007 use only
    --  the index of built-in formats and discards the strings for BIFF2, but not for BIFF3...
  begin
    --  5.12 BUILTINFMTCOUNT
    case xl.xl_format is
      when BIFF2 =>
        Write_Biff (xl, 16#001F#, Intel_16 (Unsigned_16 (last_built_in - 5)));
      when BIFF3 =>
        Write_Biff (xl, 16#0056#, Intel_16 (Unsigned_16 (last_built_in - 3)));
      when BIFF4 =>
        Write_Biff (xl, 16#0056#, Intel_16 (Unsigned_16 (last_built_in + 1)));
      --  when BIFF8 =>
      --    null;
    end case;
    --  loop & case avoid omitting any choice
    for n in Number_Format_Type'First .. last_custom_number_format loop
      case n is
        when general    =>  Write_Number_Format_String (xl, "General");
        when decimal_0  =>  Write_Number_Format_String (xl, "0");
        when decimal_2  =>  Write_Number_Format_String (xl, "0" & sep_deci & "00"); -- 'Comma' built-in style
        when decimal_0_thousands_separator =>
          Write_Number_Format_String (xl, "#" & sep_1000 & "##0");
        when decimal_2_thousands_separator =>
          Write_Number_Format_String (xl, "#" & sep_1000 & "##0" & sep_deci & "00");
        when no_currency_0       =>
          if xl.xl_format >= BIFF4 then
            Write_Number_Format_String (xl, "#" & sep_1000 & "##0;-#" & sep_1000 & "##0");
          end if;
        when no_currency_red_0   =>
          if xl.xl_format >= BIFF4 then
            Write_Number_Format_String (xl, "#" & sep_1000 & "##0;-#" & sep_1000 & "##0");
          --  [Red] doesn't go with non-English versions of Excel !!
          end if;
        when no_currency_2       =>
          if xl.xl_format >= BIFF4 then
            Write_Number_Format_String (xl,  "#" & sep_1000 & "##0" & sep_deci & "00;" &
                          "-#" & sep_1000 & "##0" & sep_deci & "00");
          end if;
        when no_currency_red_2   =>
          if xl.xl_format >= BIFF4 then
            Write_Number_Format_String (xl,  "#" & sep_1000 & "##0" & sep_deci & "00;" &
                          "-#" & sep_1000 & "##0" & sep_deci & "00");
          end if;
        when currency_0       =>
          Write_Number_Format_String (xl, "$ #" & sep_1000 & "##0;$ -#" & sep_1000 & "##0");
        when currency_red_0   =>
          Write_Number_Format_String (xl, "$ #" & sep_1000 & "##0;$ -#" & sep_1000 & "##0");
          --  [Red] doesn't go with non-English versions of Excel !!
        when currency_2       =>
          Write_Number_Format_String (xl,  "$ #" & sep_1000 & "##0" & sep_deci & "00;" &
                          "$ -#" & sep_1000 & "##0" & sep_deci & "00");
        when currency_red_2   =>
          Write_Number_Format_String (xl,  "$ #" & sep_1000 & "##0" & sep_deci & "00;" &
                          "$ -#" & sep_1000 & "##0" & sep_deci & "00");
        when percent_0        =>  Write_Number_Format_String (xl, "0%");   -- 'Percent' built-in style
        when percent_2        =>  Write_Number_Format_String (xl, "0" & sep_deci & "00%");
        when scientific       =>  Write_Number_Format_String (xl, "0" & sep_deci & "00E+00");
        when fraction_1       =>
          if xl.xl_format >= BIFF3 then
            Write_Number_Format_String (xl, "#\ ?/?");
          end if;
        when fraction_2       =>
          if xl.xl_format >= BIFF3 then
            Write_Number_Format_String (xl, "#\ ??/??");
          end if;
        when dd_mm_yyyy       =>  Write_Number_Format_String (xl, "dd/mm/yyyy");
        when dd_mmm_yy        =>  Write_Number_Format_String (xl, "dd/mmm/yy");
        when dd_mmm           =>  Write_Number_Format_String (xl, "dd/mmm");
        when mmm_yy           =>  Write_Number_Format_String (xl, "mmm/yy");
        when h_mm_AM_PM       =>  Write_Number_Format_String (xl, "h:mm\ AM/PM");
        when h_mm_ss_AM_PM    =>  Write_Number_Format_String (xl, "h:mm:ss\ AM/PM");
        when hh_mm            =>  Write_Number_Format_String (xl, "hh:mm");
        when hh_mm_ss         =>  Write_Number_Format_String (xl, "hh:mm:ss");
        when dd_mm_yyyy_hh_mm =>  Write_Number_Format_String (xl, "dd/mm/yyyy\ hh:mm");
        when percent_0_plus  =>
          Write_Number_Format_String (xl, "+0%;-0%;0%");
        when percent_2_plus  =>
          Write_Number_Format_String (xl, "+0" & sep_deci & "00%;-0" & sep_deci & "00%;0" & sep_deci & "00%");
        when date_iso        => Write_Number_Format_String (xl, "yyyy\-mm\-dd");
        when date_h_m_iso    => Write_Number_Format_String (xl, "yyyy\-mm\-dd\ hh:mm");
        when date_h_m_s_iso  => Write_Number_Format_String (xl, "yyyy\-mm\-dd\ hh:mm:ss");
          --  !! Trouble: Excel (German Excel/French locale) writes yyyy, reads it,
          --    understands it and translates it into aaaa, but is unable to
          --    understand *our* yyyy
          --  Same issue as [Red] vs [Rot] above.
      end case;
    end loop;
    --  ^ Some formats in the original list caused problems, probably
    --    because of regional placeholder symbols
    case xl.xl_format is
      when BIFF2 =>
        for i in 1 .. 6 loop
          Write_Number_Format_String (xl, "@");
        end loop;
      when BIFF3 =>
        for i in 1 .. 4 loop
          Write_Number_Format_String (xl, "@");
        end loop;
      when BIFF4 =>
        null;
    end case;
    --  ^ Stuffing for having the same number of built-in and EW custom
  end Write_Default_Number_Formats;

  --  5.35 DIMENSION
  procedure Write_Dimensions (xl : Excel_Out_Stream'Class) is
    --  sheet bounds:   0 2 Index to first used row
    --                  2 2 Index to last used row, increased by 1
    --                  4 2 Index to first used column
    --                  6 2 Index to last used column, increased by 1
    --
    --  Since our row / column counts are 1-based, no need to increase by 1.
    sheet_bounds : constant Byte_Buffer :=
      Intel_16 (0) &
      Intel_16 (Unsigned_16 (xl.maxrow)) &
      Intel_16 (0) &
      Intel_16 (Unsigned_16 (xl.maxcolumn));
    --  sheet_bounds_32_16: constant Byte_buffer:=
    --    Intel_32(0) &
    --    Intel_32(Unsigned_32(xl.maxrow)) &
    --    Intel_16(0) &
    --    Intel_16(Unsigned_16(xl.maxcolumn));
  begin
    case xl.xl_format is
      when BIFF2 =>
        Write_Biff (xl, 16#0000#, sheet_bounds);
      when BIFF3 | BIFF4 =>
        Write_Biff (xl, 16#0200#, sheet_bounds & (0, 0));
      --  when BIFF8 =>
      --    Write_Biff(xl, 16#0200#, sheet_bounds_32_16 & (0,0));
    end case;
  end Write_Dimensions;

  procedure Define_Number_Format
    (xl            : in out Excel_Out_Stream;
     format        :    out Number_Format_Type;
     format_string : in     String)
  is
  begin
    xl.number_fmt := xl.number_fmt + 1;
    case xl.xl_format is
      when BIFF2 =>
        if xl.number_fmt > 63 then
          raise Number_format_out_of_range
            with "Only 64 number formats are allowed in the BIFF2 format";
            --  Reason: see encoding in the Cell Attributes (2.5.13).
        end if;
      when BIFF3 .. BIFF4 =>
        if xl.number_fmt > 255 then
          raise Number_format_out_of_range
            with "Only 256 number formats are allowed in the BIFF3, BIFF4 formats";
        end if;
    end case;
    format := xl.number_fmt;
    Write_Number_Format_String (xl, format_string);
  end Define_Number_Format;

  procedure Write_Worksheet_header (xl : in out Excel_Out_Stream'Class) is

    procedure Define_Style (fmt : Format_Type; style_id : Unsigned_8) is
      Base_Level : constant := 255;
    begin
      Write_Biff (xl,
        16#0293#,
        Intel_16 (Unsigned_16 (fmt) + 16#8000#) & style_id & Base_Level
      );
    end Define_Style;
    --
    Comma_Style     : constant := 3;
    Currency_Style  : constant := 4;
    Percent_Style   : constant := 5;
    font_for_styles, font_2, font_3 : Font_Type;
    --
    function Encoding_code return Unsigned_16 is  --  5.17 CODEPAGE, p. 145
    begin
      case xl.encoding is
        when Windows_CP_874  => return 874;
        when Windows_CP_932  => return 932;
        when Windows_CP_936  => return 936;
        when Windows_CP_949  => return 949;
        when Windows_CP_950  => return 950;
        when Windows_CP_1250 => return 1250;
        when Windows_CP_1251 => return 1251;
        when Windows_CP_1252 =>
          case xl.xl_format is
            when BIFF2 .. BIFF3 =>
              return 16#8001#;
            when BIFF4 =>
              return 1252;
          end case;
        when Windows_CP_1253 => return 1253;
        when Windows_CP_1254 => return 1254;
        when Windows_CP_1255 => return 1255;
        when Windows_CP_1256 => return 1256;
        when Windows_CP_1257 => return 1257;
        when Windows_CP_1258 => return 1258;
        when Windows_CP_1361 => return 1361;
        when Apple_Roman     => return 10000;
      end case;
    end Encoding_code;
    --
  begin
    Write_BOF (xl);
    --  5.17 CODEPAGE, p. 145
    case xl.xl_format is
      --  when BIFF8 =>   --  UTF-16
      --    Write_Biff(xl, 16#0042#, Intel_16(16#04B0#));
      when others =>
        Write_Biff (xl, 16#0042#, Intel_16 (Encoding_code));
    end case;
    --  5.14 CALCMODE
    Write_Biff (xl, 16#000D#, Intel_16 (1)); --  1 => automatic
    --  5.85 REFMODE
    Write_Biff (xl, 16#000F#, Intel_16 (1)); --  1 => A1 mode
    --  5.28 DATEMODE
    Write_Biff (xl, 16#0022#, Intel_16 (0)); --  0 => 1900; 1 => 1904 Date system
    --  NB: the 1904 variant (Mac) is ignored by LibreOffice (<= 3.5), then wrong dates !
    --
    Define_Font (xl, "Arial",   10, xl.def_font);
    Define_Font (xl, "Arial",   10, font_for_styles); -- Used by BIFF3+'s styles
    Define_Font (xl, "Calibri", 10, font_2); -- Defined in BIFF3 files written by Excel 2002
    Define_Font (xl, "Calibri", 10, font_3); -- Defined in BIFF3 files written by Excel 2002
    Write_Default_Number_Formats (xl);
    --  5.111 WINDOWPROTECT
    Write_Biff (xl, 16#0019#, Intel_16 (0));
    --  Define default format
    Define_Format (xl, xl.def_font, general, xl.def_fmt);
    if xl.xl_format >= BIFF3 then
      --  Don't ask why we need the following useless formats, but it is as Excel 2002
      --  write formats. Additionally, the default format is turned into decimal_2
      --  when a file without those useless formats is opened in Excel (2002) !
      Define_Format (xl, font_for_styles, general, xl.def_fmt);
      Define_Format (xl, font_for_styles, general, xl.def_fmt);
      Define_Format (xl, font_2, general, xl.def_fmt);
      Define_Format (xl, font_2, general, xl.def_fmt);
      for i in 5 .. 15 loop
        Define_Format (xl, xl.def_font, general, xl.def_fmt);
      end loop;
      --  Final default format index is the last changed xl.def_fmt
    end if;
    Use_Default_Format (xl);
    --  Define formats for the BIFF3+ "styles":
    Define_Format (xl, font_for_styles, decimal_2, xl.cma_fmt);
    Define_Format (xl, font_for_styles, currency_0, xl.ccy_fmt);
    Define_Format (xl, font_for_styles, percent_0, xl.pct_fmt);
    --  Define styles - 5.103 STYLE p. 212
    --  NB: - it is BIFF3+ (we cheat a bit if selected format is BIFF2).
    --      - these "styles" seem to be a zombie feature of Excel 3
    --      - the whole purpose of including this is because format
    --        buttons (%)(,) in Excel 95 through 2007 are using these styles;
    --        if the styles are not defined, those buttons are not working
    --        when an Excel Writer sheet is open in MS Excel.
    Define_Style (xl.cma_fmt, Comma_Style);
    Define_Style (xl.ccy_fmt, Currency_Style);
    Define_Style (xl.pct_fmt, Percent_Style);
    xl.dimrecpos := Index (xl);
    Write_Dimensions (xl);
    xl.is_created := True;
  end Write_Worksheet_header;

  type Font_or_Background is (for_font, for_background);
  type Color_pair is array (Font_or_Background) of Unsigned_16;
  auto_color : constant Color_pair :=
    (16#7FFF#, -- system window text colour
     16#0019#  -- system window background colour
    );

  color_code : constant array (Excel_Type, Color_Type) of Color_pair :=
    (BIFF2 =>
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
      BIFF3 | BIFF4 =>
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

  --  *** Exported procedures **********************************************

  --  5.115 XF - Extended Format
  procedure Define_Format
    (xl               : in out Excel_Out_Stream;
     font             : in     Font_Type;           --  Default_font(xl), or given by Define_font
     number_format    : in     Number_Format_Type;  --  built-in, or given by Define_number_format
     cell_format      :    out Format_Type;
     -- Optional parameters --
     horizontal_align : in     Horizontal_Alignment := general_alignment;
     border           : in     Cell_Border          := no_border;
     shaded           : in     Boolean              := False;    --  Add a dotted background pattern
     background_color : in     Color_Type           := automatic;
     wrap_text        : in     Boolean              := False;
     vertical_align   : in     Vertical_Alignment   := bottom_alignment;
     text_orient      : in     Text_Orientation     := normal)
  is
    actual_number_format : Number_Format_Type := number_format;
    cell_is_locked : constant := 1;
    --  ^ Means actually: cell formula protection is possible, and enabled when sheet is protected.
    procedure Define_BIFF2_XF is
      border_bits, mask : Unsigned_8;
    begin
      border_bits := 0;
      mask := 8;
      for s in Cell_Border_Single loop
        if border (s) then
          border_bits := border_bits + mask;
        end if;
        mask := mask * 2;
      end loop;
      --  5.115.2 XF Record Contents, p. 221 for BIFF2
      Write_Biff (
        xl,
        16#0043#, -- XF code in BIFF2
        (Unsigned_8 (font),
         --  ^ Index to FONT record
         0,
         --  ^ Not used
         Number_Format_Type'Pos (actual_number_format) + 16#40# * cell_is_locked,
         --  ^ Number format and cell flags
         Horizontal_Alignment'Pos (horizontal_align) +
         border_bits +
         Boolean'Pos (shaded) * 128
         --  ^ Horizontal alignment, border style, and background
        )
      );
    end Define_BIFF2_XF;

    area_code : Unsigned_16;

    procedure Define_BIFF3_XF is
    begin
      --  5.115.2 XF Record Contents, p. 221 for BIFF3
      Write_Biff (
        xl,
        16#0243#, -- XF code in BIFF3
        (Unsigned_8 (font),
         --  ^ 0 - Index to FONT record
         Number_Format_Type'Pos (actual_number_format),
         --  ^ 1 - Number format and cell flags
         cell_is_locked,
         --  ^ 2 - XF_TYPE_PROT (5.115.1)
         16#FF#
         --  ^ 3 - XF_USED_ATTRIB
        ) &
        Intel_16 (
          Horizontal_Alignment'Pos (horizontal_align) +
          Boolean'Pos (wrap_text) * 8
        ) &
        --  ^ 4 - Horizontal alignment, text break, parent style XF
        Intel_16 (area_code) &
        --  ^ 6 - XF_AREA_34
        (Boolean'Pos (border (top_single)),
           Boolean'Pos (border (left_single)),
           Boolean'Pos (border (bottom_single)),
           Boolean'Pos (border (right_single))
        )
        --  ^ 8 - XF_BORDER_34 - thin (=1) line; we could have other line styles:
        --       Thin, Medium, Dashed, Dotted, Thick, Double, Hair
      );
    end Define_BIFF3_XF;

    procedure Define_BIFF4_XF is
    begin
      --  5.115.2 XF Record Contents, p. 222 for BIFF4
      Write_Biff (
        xl,
        16#0443#, -- XF code in BIFF4
        (Unsigned_8 (font),
         --  ^ 0 - Index to FONT record
         Number_Format_Type'Pos (actual_number_format),
         --  ^ 1 - Number format and cell flags
         cell_is_locked, 0,
         --  ^ 2 - XF type, cell protection, and parent style XF
         Horizontal_Alignment'Pos (horizontal_align) +
         Boolean'Pos (wrap_text) * 8 +
         (Vertical_Alignment'Pos (vertical_align) and 3) * 16 +
         Text_Orientation'Pos (text_orient) * 64,
         --  ^ 4 - Alignment (hor & ver), text break, and text orientation
         16#FF#
         --  ^ 3 - XF_USED_ATTRIB
        ) &
        --  ^ 4 - Horizontal alignment, text break, parent style XF
        Intel_16 (area_code) &
        --  ^ 6 - XF_AREA_34
        (Boolean'Pos (border (top_single)),
           Boolean'Pos (border (left_single)),
           Boolean'Pos (border (bottom_single)),
           Boolean'Pos (border (right_single))
        )
        --  ^ 8 - XF_BORDER_34 - thin (=1) line; we could have other line styles:
        --        Thin, Medium, Dashed, Dotted, Thick, Double, Hair
      );
    end Define_BIFF4_XF;

  begin
    --  2.5.12 Patterns for Cell and Chart Background Area
    --  This is for BIFF3+
    if shaded then
      area_code :=
        Boolean'Pos (shaded) * 17 +                        -- Sparse pattern, like BIFF2 "shade"
        16#40#  * color_code (BIFF3, black)(for_background) +           -- pattern colour
        16#800# * color_code (BIFF3, background_color)(for_background); -- pattern background
    elsif background_color = automatic then
      area_code := 0;
    else
      area_code :=
        1 +                                                          -- Full pattern
        16#40#  * color_code (BIFF3, background_color)(for_background) +  -- pattern colour
        16#800# * color_code (BIFF3, background_color)(for_background);   -- pattern background
    end if;
    case xl.xl_format is
      when BIFF2 =>
        case actual_number_format is
          when general .. no_currency_2 =>
            null;
          when currency_0 .. fraction_2 =>
            actual_number_format := actual_number_format - 4;
          when dd_mm_yyyy .. last_custom_number_format =>
            actual_number_format := actual_number_format - 6;
          when others =>
            null;
        end case;
        Define_BIFF2_XF;
      when BIFF3 =>
        if actual_number_format in currency_0 .. last_custom_number_format then
          actual_number_format := actual_number_format - 4;
        end if;
        Define_BIFF3_XF;
      when BIFF4 =>
        Define_BIFF4_XF;
      --  when BIFF8 =>
      --    Define_BIFF8_XF;  --  BIFF8: 16#00E0#, p. 224
    end case;
    --
    --  Now we will store the newly defined format.
    --
    xl.xfs := xl.xfs + 1;
    if xl.xfs not in XF_Range then
      raise Format_out_of_range
        with
          "Too many formats defined, maximum number of formats " &
          "(including a few pre-defined) is" &
          Integer'Image (XF_Range'Last + 1);
    end if;
    cell_format := Format_Type (xl.xfs);
    xl.xf_def (xl.xfs) := (font => font, numb => number_format);
  end Define_Format;

  procedure Header (xl : Excel_Out_Stream; page_header_string : String) is
  begin
    Write_Biff (xl, 16#0014#, To_buf_8_bit_length (page_header_string)); -- 5.55 p.180
  end Header;

  procedure Footer (xl : Excel_Out_Stream; page_footer_string : String) is
  begin
    Write_Biff (xl, 16#0015#, To_buf_8_bit_length (page_footer_string)); -- 5.48 p.173
  end Footer;

  procedure Left_Margin (xl : Excel_Out_Stream; inches : Long_Float) is
  begin
    Write_Biff (xl, 16#0026#, IEEE_Double_Intel (inches));
  end Left_Margin;

  procedure Right_Margin (xl : Excel_Out_Stream; inches : Long_Float) is
  begin
    Write_Biff (xl, 16#0027#, IEEE_Double_Intel (inches));
  end Right_Margin;

  procedure Top_Margin (xl : Excel_Out_Stream; inches : Long_Float) is
  begin
    Write_Biff (xl, 16#0028#, IEEE_Double_Intel (inches));
  end Top_Margin;

  procedure Bottom_Margin (xl : Excel_Out_Stream; inches : Long_Float) is
  begin
    Write_Biff (xl, 16#0029#, IEEE_Double_Intel (inches));
  end Bottom_Margin;

  procedure Margins
    (xl : Excel_Out_Stream;
     left_inches,
     right_inches,
     top_inches,
     bottom_inches : Long_Float)
  is
  begin
    Left_Margin (xl, left_inches);
    Right_Margin (xl, right_inches);
    Top_Margin (xl, top_inches);
    Bottom_Margin (xl, bottom_inches);
  end Margins;

  procedure Print_Row_Column_Headers (xl : Excel_Out_Stream) is
  begin
    Write_Biff (xl, 16#002A#, Intel_16 (1));  --  5.81 PRINTHEADERS p.199
  end  Print_Row_Column_Headers;

  procedure Print_Gridlines (xl : Excel_Out_Stream) is
  begin
    Write_Biff (xl, 16#002B#, Intel_16 (1));  --  5.80 PRINTGRIDLINES p.199
  end Print_Gridlines;

  procedure Page_Setup (
    xl                     : Excel_Out_Stream;
    scaling_percents       : Positive := 100;
    fit_width_with_n_pages : Natural := 1; -- 0: as many as possible
    fit_height_with_n_pages : Natural := 1; -- 0: as many as possible
    orientation            : Orientation_Choice := portrait;
    scale_or_fit           : Scale_or_Fit_Choice := scale
  )
  is
  begin
    --  5.73 PAGESETUP p.192 - this is BIFF4+ (cheat if xl.format below)!
    Write_Biff (xl,
      16#00A1#,
      Intel_16 (0) & -- paper type undefined
      Intel_16 (Unsigned_16 (scaling_percents)) &
      Intel_16 (1) & -- start page number
      Intel_16 (Unsigned_16 (fit_width_with_n_pages)) &
      Intel_16 (Unsigned_16 (fit_height_with_n_pages)) &
      Intel_16 (2 * Orientation_Choice'Pos (orientation))
    );
    --  5.97 SHEETPR p.207 - this is BIFF3+ (cheat if xl.format below) !
    --  NB: this field contains other informations, should be delayed
    --        in case other preferences are to be set
    Write_Biff (xl,
      16#0081#,
      Intel_16 (256 * Scale_or_Fit_Choice'Pos (scale_or_fit))
    );
  end Page_Setup;

  y_scale : constant := 20; -- scaling to obtain character point (pt) units

  --  5.31 DEFAULTROWHEIGHT
  procedure Write_Default_Row_Height (
        xl     : Excel_Out_Stream;
        height : Positive
  )
  is
    default_twips : constant Byte_Buffer := Intel_16 (Unsigned_16 (height * y_scale));
    options_flags : constant Byte_Buffer := (1, 0);
    --  1 = Row height and default font height do not match
  begin
    case xl.xl_format is
      when BIFF2 =>
        Write_Biff (xl, 16#0025#, default_twips);
      when BIFF3 | BIFF4 =>
        Write_Biff (xl, 16#0225#, options_flags & default_twips);
    end case;
  end Write_Default_Row_Height;

  --  5.32 DEFCOLWIDTH
  procedure Write_Default_Column_Width (
        xl : in out Excel_Out_Stream;
        width  : Positive)
  is
  begin
    Write_Biff (xl, 16#0055#, Intel_16 (Unsigned_16 (width)));
    xl.defcolwdth := 256 * width;
  end Write_Default_Column_Width;

  procedure Write_Column_Width (
        xl     : in out Excel_Out_Stream;
        column : Positive;
        width  : Natural)
  is
  begin
    Write_Column_Width (xl, column, column, width);
  end Write_Column_Width;

  procedure Write_Column_Width (
    xl            : in out Excel_Out_Stream;
    first_column,
    last_column   : Positive;
    width         : Natural
  )
  is
  begin
    case xl.xl_format is
      when BIFF2 =>
        --  5.20 COLWIDTH (BIFF2 only)
        Write_Biff (xl, 16#0024#,
          Unsigned_8 (first_column - 1) &
          Unsigned_8 (last_column - 1) &
          Intel_16 (Unsigned_16 (width * 256)));
      when BIFF3 | BIFF4 =>
        --  5.18 COLINFO (BIFF3+)
        Write_Biff (xl, 16#007D#,
          Intel_16 (Unsigned_16 (first_column - 1)) &
          Intel_16 (Unsigned_16 (last_column - 1)) &
          Intel_16 (Unsigned_16 (width * 256)) &
          Intel_16 (0) & -- Index to XF record (5.115) for default column formatting
          Intel_16 (0) & -- Option flags
          (0, 0)         -- Not used
        );
        for j in first_column .. last_column loop
          xl.std_col_width (j) := False;
        end loop;
    end case;
  end Write_Column_Width;

  --  5.88 ROW
  --  The OpenOffice documentation tells nice stories about row blocks,
  --  but single ROW commands can also be put before in the data stream,
  --  where the column widths are set. Excel saves with blocks of ROW
  --  commands, most of them useless.

  procedure Write_Row_Height (
    xl     : Excel_Out_Stream;
    row    : Positive;
    height : Natural
  )
  is
    row_info_base : Byte_Buffer :=
      Intel_16 (Unsigned_16 (row - 1)) &
      Intel_16 (0)   & -- col. min.
      Intel_16 (255) & -- col. max.
      Intel_16 (Unsigned_16 (height * y_scale));
    fDyZero : Unsigned_8 := 0;
  begin
    case xl.xl_format is
      when BIFF2 =>
        Write_Biff (xl, 16#0008#,
          row_info_base &
          (1 .. 3 => 0) &
          Intel_16 (0) -- offset to data
        );
      when BIFF3 | BIFF4 =>
        if height = 0 then -- proper hiding (needed with LibreOffice)
          fDyZero := 1;
          row_info_base (row_info_base'Last - 1 .. row_info_base'Last) :=
            Intel_16 (16#8000#);
        end if;
        Write_Biff (xl, 16#0208#,
          row_info_base &
          --  http://msdn.microsoft.com/en-us/library/dd906757(v=office.12).aspx
          (0, 0,  -- reserved1 (2 bytes): MUST be zero, and MUST be ignored.
           0, 0,  -- unused1 (2 bytes): Undefined and MUST be ignored.
           fDyZero *  32 +  -- D - fDyZero (1 bit): row is hidden
                 1 *  64 +  -- E - fUnsynced (1 bit): row height was manually set
                 0 * 128,   -- F - fGhostDirty (1 bit): the row was formatted
           1) &   -- reserved3 (1 byte): MUST be 1, and MUST be ignored
           Intel_16 (15)
           --  ^ ixfe_val, then 4 bits.
           --    If fGhostDirty is 0, ixfe_val is undefined and MUST be ignored.
        );
    end case;
  end Write_Row_Height;

  --  5.45 FONT, p.171
  procedure Define_Font
    (xl           : in out Excel_Out_Stream;
     font_name    :        String;
     height       :        Positive;
     font         :    out Font_Type;
     style        :        Font_Style := regular;
     color        :        Color_Type := automatic)
  is
    style_bits, mask : Unsigned_16;
  begin
    style_bits := 0;
    mask := 1;
    for s in Font_Style_Single loop
      if style (s) then
        style_bits := style_bits + mask;
      end if;
      mask := mask * 2;
    end loop;
    xl.fonts := xl.fonts + 1;
    if xl.fonts = 4 then
      xl.fonts := 5;
      --  Anomaly! The font with index 4 is omitted in all BIFF versions (5.45).
      --  Numbering is 0, 1, 2, 3, *5*, 6,...
    end if;
    case xl.xl_format is
      when BIFF2 =>
        if xl.fonts > 3 then
          raise Font_out_of_range with "Only 4 fonts are allowed in the BIFF2 format";
          --  Reason: in the Cell Attributes (2.5.13), font index is encoded on 2 bits!
        end if;
        Write_Biff (xl, 16#0031#,
          Intel_16 (Unsigned_16 (height * y_scale)) &
          Intel_16 (style_bits) &
          To_buf_8_bit_length (font_name)
        );
        if color /= automatic then
          --  5.47 FONTCOLOR
          Write_Biff (xl, 16#0045#, Intel_16 (color_code (BIFF2, color)(for_font)));
        end if;
      when BIFF3 | BIFF4 =>  --  BIFF8 has 16#0031#, p. 171
        if xl.fonts > 255 then
          raise Font_out_of_range with "Only 256 fonts are allowed in the BIFF3, BIFF4 formats";
        end if;
        Write_Biff (xl, 16#0231#,
          Intel_16 (Unsigned_16 (height * y_scale)) &
          Intel_16 (style_bits) &
          Intel_16 (color_code (BIFF3, color)(for_font)) &
          To_buf_8_bit_length (font_name)
        );
    end case;
    font := Font_Type (xl.fonts);
  end Define_Font;

  procedure Jump_to_and_store_max (xl : in out Excel_Out_Stream; r, c : Integer) is
    pragma Inline (Jump_to_and_store_max);
  begin
    if not xl.is_created then
      raise Excel_stream_not_created;
    end if;
    Jump_to (xl, r, c); -- Store and check current position
    if r > xl.maxrow then
      xl.maxrow := r;
    end if;
    if c > xl.maxcolumn then
      xl.maxcolumn := c;
    end if;
  end Jump_to_and_store_max;

  --  2.5.13 Cell Attributes (BIFF2 only)
  function Cell_attributes (xl : Excel_Out_Stream) return Byte_Buffer is
  begin
    return
      (Unsigned_8 (xl.xf_in_use),
       Unsigned_8 (xl.xf_def (xl.xf_in_use).numb) + 16#40# *
       Unsigned_8 (xl.xf_def (xl.xf_in_use).font),
       0
      );
  end Cell_attributes;

  function Almost_zero (x : Long_Float) return Boolean is
  begin
    return abs x <= Long_Float'Model_Small;
  end Almost_zero;

  --  5.50 FORMULA, p.176 - Internal
  --
  procedure Write_as_Formula
    (xl                    : in out Excel_Out_Stream;
     r, c                  :        Positive;
     str                   :        String;
     formula_result        :        Long_Float;
     recalculate_on_change :        Boolean;
     recalculate_on_open   :        Boolean;
     is_valid              :    out Boolean)
  is

    --  3.1 Common Formula Structure, p.28
    --
    function Tokenize return Byte_Buffer is
      --  3.4.2 Binary Operator Tokens, p.40
      tAdd : constant := 16#03#;
      tMul : constant := 16#05#;
      --  3.4.4 Constant Operand Tokens, p.41
      tInt : constant := 16#1E#;  --  3.8.5
      --  Example of page 30 (2*4+5):
      test_data : constant Byte_Buffer :=
        tInt & Intel_16 (2) &
        tInt & Intel_16 (4) &
        tMul &
        tInt & Intel_16 (5) &
        tAdd;
    begin
      is_valid := True;
      return test_data;
    end Tokenize;

    function Tokenize_with_Size return Byte_Buffer is
      data : constant Byte_Buffer := Tokenize;
    begin
      case xl.xl_format is
        when BIFF2 =>
          if data'Length > 255 then
            is_valid := False;
            return empty_buffer;
          end if;
          return data'Length & data;
        when BIFF3 | BIFF4 =>
          if data'Length > 65535 then
            is_valid := False;
            return empty_buffer;
          end if;
          return Intel_16 (data'Length) & data;
      end case;
    end Tokenize_with_Size;

    rpn_token_array : constant Byte_Buffer := Tokenize_with_Size;
    --  3.1.6 Additional Token Data:
    additional_token_data : constant Byte_Buffer := empty_buffer;
  begin
    if is_valid then
      Jump_to_and_store_max (xl, r, c);
      case xl.xl_format is
        when BIFF2 =>
          Write_Biff
            (xl, 16#0006#,
             Intel_16 (Unsigned_16 (r - 1)) &
             Intel_16 (Unsigned_16 (c - 1)) &
             Cell_attributes (xl) &
             IEEE_Double_Intel (formula_result) &
             Boolean'Pos (recalculate_on_change) &
             rpn_token_array);
        when BIFF3 | BIFF4 =>
          Write_Biff
            (xl, (if xl.xl_format = BIFF3 then 16#0206# else 16#0406#),
             Intel_16 (Unsigned_16 (r - 1)) &
             Intel_16 (Unsigned_16 (c - 1)) &
             Intel_16 (Unsigned_16 (xl.xf_in_use)) &
             IEEE_Double_Intel (formula_result) &
             Intel_16
                  (Boolean'Pos (recalculate_on_change) +
               2 * Boolean'Pos (recalculate_on_open)) &
             rpn_token_array);
      end case;
      Jump_to (xl, r, c + 1);  --  Store and check new position
    end if;
  end Write_as_Formula;

  --  5.71 NUMBER - Internal
  --
  procedure Write_as_Double
    (xl   : in out Excel_Out_Stream;
     r, c :        Positive;
     num  :        Long_Float)
  is
    pragma Inline (Write_as_Double);
  begin
    Jump_to_and_store_max (xl, r, c);
    case xl.xl_format is
      when BIFF2 =>
        Write_Biff
          (xl, 16#0003#,
           Intel_16 (Unsigned_16 (r - 1)) &
           Intel_16 (Unsigned_16 (c - 1)) &
           Cell_attributes (xl) &
           IEEE_Double_Intel (num));
      when BIFF3 | BIFF4 =>
        Write_Biff
          (xl, 16#0203#,
           Intel_16 (Unsigned_16 (r - 1)) &
           Intel_16 (Unsigned_16 (c - 1)) &
           Intel_16 (Unsigned_16 (xl.xf_in_use)) &
           IEEE_Double_Intel (num));
    end case;
    Jump_to (xl, r, c + 1);  --  Store and check new position
  end Write_as_Double;

  --  Internal. This is BIFF2 only. BIFF format choice unchecked here.
  --
  procedure Write_as_16_bit_unsigned
    (xl   : in out Excel_Out_Stream;
     r, c :        Positive;
     num  :        Unsigned_16)
  is
    pragma Inline (Write_as_16_bit_unsigned);
  begin
    Jump_to_and_store_max (xl, r, c);
    --  5.60 INTEGER
    Write_Biff (xl, 16#0002#,
      Intel_16 (Unsigned_16 (r - 1)) &
      Intel_16 (Unsigned_16 (c - 1)) &
      Cell_attributes (xl) &
      Intel_16 (num)
    );
    Jump_to (xl, r, c + 1); -- Store and check new position
  end Write_as_16_bit_unsigned;

  --  Internal. This is BIFF3+. BIFF format choice unchecked here.
  --
  procedure Write_as_30_bit_signed
    (xl   : in out Excel_Out_Stream;
     r, c :        Positive;
     num  :        Integer_32)
  is
    pragma Inline (Write_as_30_bit_signed);
    RK_val : Unsigned_32;
    RK_code : constant := 2; -- Code for signed integer. See 2.5.5 RK Values
  begin
    if num >= 0 then
      RK_val := Unsigned_32 (num) * 4 + RK_code;
    else
      RK_val := (-Unsigned_32 (-num)) * 4 + RK_code;
    end if;
    Jump_to_and_store_max (xl, r, c);
    --  5.87 RK
    Write_Biff (xl, 16#027E#,
      Intel_16 (Unsigned_16 (r - 1)) &
      Intel_16 (Unsigned_16 (c - 1)) &
      Intel_16 (Unsigned_16 (xl.xf_in_use)) &
      Intel_32 (RK_val)
    );
    Jump_to (xl, r, c + 1); -- Store and check new position
  end Write_as_30_bit_signed;

  --
  --  Profile with floating-point number
  --
  procedure Write
    (xl   : in out Excel_Out_Stream;
     r, c :        Positive;
     num  :        Long_Float)
  is
    max_16_u : constant :=   2.0 ** 16 - 1.0;
    min_30_s : constant := -(2.0 ** 29);
    max_30_s : constant :=   2.0 ** 29 - 1.0;
  begin
    case xl.xl_format is
      when BIFF2 =>
        if num in 0.0 .. max_16_u and then
           Almost_zero (num - Long_Float'Floor (num))
        then
          Write_as_16_bit_unsigned (xl, r, c, Unsigned_16 (Long_Float'Floor (num)));
        else
          Write_as_Double (xl, r, c, num);
        end if;
      when BIFF3 | BIFF4 =>
        if num in min_30_s .. max_30_s and then
           Almost_zero (num - Long_Float'Floor (num))
        then
          Write_as_30_bit_signed (xl, r, c, Integer_32 (Long_Float'Floor (num)));
        else
          Write_as_Double (xl, r, c, num);
        end if;
    end case;
  end Write;

  --
  --  Profile with integer number
  --
  procedure Write
    (xl   : in out Excel_Out_Stream;
     r, c :        Positive;
     num  :        Integer)
  is
  begin
    --  We use an integer representation (and small storage) if possible;
    --  we need to use a floating-point in all other cases
    case xl.xl_format is
      when BIFF2 =>
        if num in 0 .. 2**16 - 1 then
          Write_as_16_bit_unsigned (xl, r, c, Unsigned_16 (num));
        else
          Write_as_Double (xl, r, c, Long_Float (num));
        end if;
      when BIFF3 | BIFF4 =>
        if num in -2**29 .. 2**29 - 1 then
          Write_as_30_bit_signed (xl, r, c, Integer_32 (num));
        else
          Write_as_Double (xl, r, c, Long_Float (num));
        end if;
    end case;
  end Write;

  --  --  Function taken from Wasabee.Encoding.
  --  function ISO_8859_1_to_UTF_16(s: String) return Wide_String is
  --    --  This conversion is a trivial 8-bit to 16-bit copy.
  --    r: Wide_String(s'Range);
  --  begin
  --    for i in s'Range loop
  --      r(i):= Wide_Character'Val(Character'Pos(s(i)));
  --    end loop;
  --    return r;
  --  end ISO_8859_1_to_UTF_16;

  --  5.63 LABEL
  --
  procedure Write_as_Label
    (xl   : in out Excel_Out_Stream;
     r, c :        Positive;
     str  :        String)
  is
  begin
    Jump_to_and_store_max (xl, r, c);
    if str'Length > 0 then
      case xl.xl_format is
        when BIFF2 =>
          Write_Biff (xl, 16#0004#,
            Intel_16 (Unsigned_16 (r - 1)) &
            Intel_16 (Unsigned_16 (c - 1)) &
            Cell_attributes (xl) &
            To_buf_8_bit_length (str)
          );
        when BIFF3 | BIFF4 =>
          Write_Biff (xl, 16#0204#,
            Intel_16 (Unsigned_16 (r - 1)) &
            Intel_16 (Unsigned_16 (c - 1)) &
            Intel_16 (Unsigned_16 (xl.xf_in_use)) &
            To_buf_16_bit_length (str)
          );
        --  when BIFF8 =>
        --    Write_Biff(xl, 16#0204#,
        --      Intel_16(Unsigned_16(r-1)) &
        --      Intel_16(Unsigned_16(c-1)) &
        --      Intel_16(Unsigned_16(xl.xf_in_use)) &
        --      To_buf_16_bit_length(ISO_8859_1_to_UTF_16(str))
        --    );
      end case;
    end if;
    Jump_to (xl, r, c + 1);  --  Store and check new position
  end Write_as_Label;

  procedure Write
    (xl   : in out Excel_Out_Stream;
     r, c :        Positive;
     str  :        String)
  is
    is_valid : Boolean;
  begin
    --  !! Uncomment following lines when the formula parser is OK !!
    --  if str'Length > 0 and str (str'First) in '+' | '-' | '=' then
    --    Write_as_Formula (xl, r, c, str, 0.0, True, True, is_valid);
    --    if is_valid then
    --      return;  --  Formula is valid and has been written.
    --    end if;
    --  end if;
    Write_as_Label (xl, r, c, str);
  end Write;

  procedure Write (xl : in out Excel_Out_Stream; r, c : Positive; str : Ada.Strings.Unbounded.Unbounded_String)
  is
  begin
    Write (xl, r, c, To_String (str));
  end Write;

  --  Excel uses a floating-point type for time - ouch!
  --
  function To_Number (date : Ada.Calendar.Time) return Long_Float is
    --  1901 is the lowest year supported by Ada.Calendar.
    --  1900 is not a leap year, but Lotus 1-2-3, then Excel, consider it
    --  as a leap year. So, with 1901, we skip that issue anyway...
    --
    function Days_since_1901 (y, m, d : Integer) return Integer is
      function Is_leap (y : Integer) return Boolean is
      begin
        if y mod 4 = 0 then
          if y mod 100 = 0 then
            if y mod 400 = 0 then
              return True;
            else
              return False;
            end if;
          else
            return True;
          end if;
        else
          return False;
        end if;
      end Is_leap;
      days_of_previous_months : Integer;
      days_of_previous_years  : Integer;
      y_diff, y_diff_4, y_diff_100, y_diff_400 : Integer;
    begin
      case m is
        when 02 => days_of_previous_months := 31;
        when 03 => days_of_previous_months := 59;
        when 04 => days_of_previous_months := 90;
        when 05 => days_of_previous_months := 120;
        when 06 => days_of_previous_months := 151;
        when 07 => days_of_previous_months := 181;
        when 08 => days_of_previous_months := 212;
        when 09 => days_of_previous_months := 243;
        when 10 => days_of_previous_months := 273;
        when 11 => days_of_previous_months := 304;
        when 12 => days_of_previous_months := 334;
        when others => days_of_previous_months := 0;
      end case;
      if m > 2 and then Is_leap (y) then  --  February has 29 days in leap years.
        days_of_previous_months := days_of_previous_months + 1;
      end if;
      --
      y_diff     := (y - 1)       - 1900;
      y_diff_4   := (y - 1) / 4   - 1900 / 4;
      y_diff_100 := (y - 1) / 100 - 1900 / 100;
      y_diff_400 := (y - 1) / 400 - 1900 / 400;
      --  Add extra days of leap years from 1901 (included) to year y (excluded).
      days_of_previous_years := 365 * y_diff + y_diff_4 - y_diff_100 + y_diff_400;
      --
      return days_of_previous_years + days_of_previous_months + d - 1;
    end Days_since_1901;
    --
    use Ada.Calendar;
    sec : constant Day_Duration := Seconds (date);
  begin
    --  With GNAT and perhaps other systems, Duration's range allows the following:
    --    return Long_Float(date - Time_Of(1901, 01, 01, 0.0)) / 86_400.0 + 367.0;
    --  With ObjectAda and perhaps other systems, we need to count days since 1900 ourselves.
    return
      Long_Float (sec) / 86_400.0 +
      Long_Float (Days_since_1901 (Year (date), Month (date), Day (date))) +
      367.0;  --  Days from 1899-12-31 to 1901-01-01.
      --  Lotus 1-2-3, then Excel, are based on 1899-12-31 (and believe it is 1900-01-01).
  end To_Number;

  procedure Write (xl : in out Excel_Out_Stream; r, c : Positive; date : Ada.Calendar.Time)
  is
  begin
    Write (xl, r, c, To_Number (date));
  end Write;

  --  Ada.Text_IO - like. No need to specify row & column each time
  procedure Put (xl : in out Excel_Out_Stream; num : Long_Float) is
  begin
    Write (xl, xl.curr_row, xl.curr_col, num);
  end Put;

  procedure Put (xl    : in out Excel_Out_Stream;
                num   : in Integer;
                width : in Ada.Text_IO.Field := 0; -- ignored
                base  : in Ada.Text_IO.Number_Base := 10
            )
  is
  begin
    if base = 10 then
      Write (xl, xl.curr_row, xl.curr_col, num);
    else
      declare
        use Ada.Strings.Fixed;
        s : String (1 .. 50 + 0 * width);
        --  0*width is just to skip a warning of width being unused
        package IIO is new Ada.Text_IO.Integer_IO (Integer);
      begin
        IIO.Put (s, num, Base => base);
        Put (xl, Trim (s, Ada.Strings.Left));
      end;
    end if;
  end Put;

  procedure Put (xl : in out Excel_Out_Stream; str : String) is
  begin
    Write (xl, xl.curr_row, xl.curr_col, str);
  end Put;

  procedure Put (xl : in out Excel_Out_Stream; str : Ada.Strings.Unbounded.Unbounded_String) is
  begin
    Put (xl, To_String (str));
  end Put;

  procedure Put (xl : in out Excel_Out_Stream; date : Ada.Calendar.Time) is
  begin
    Put (xl, To_Number (date));
  end Put;

  procedure Merge (xl : in out Excel_Out_Stream; cells : Positive) is

    --  5.7 BLANK
    procedure Blank (r, c : Positive) is
    begin
      Jump_to_and_store_max (xl, r, c);
      case xl.xl_format is
        --  NB: Only with BIFF4, and only OpenOffice
        --  considers the cells really merged.
        when BIFF2 =>
          Write_Biff (xl, 16#0001#,
            Intel_16 (Unsigned_16 (r - 1)) &
            Intel_16 (Unsigned_16 (c - 1)) &
            Cell_attributes (xl)
          );
        when BIFF3 | BIFF4 =>
          Write_Biff (xl, 16#0201#,
            Intel_16 (Unsigned_16 (r - 1)) &
            Intel_16 (Unsigned_16 (c - 1)) &
            Intel_16 (Unsigned_16 (xl.xf_in_use))
          );
      end case;
      Jump_to (xl, r, c + 1); -- Store and check new position
    end Blank;
  begin
    for i in 1 .. cells loop
      Blank (xl.curr_row, xl.curr_col);
    end loop;
  end Merge;

  procedure Write_Cell_Comment (xl : Excel_Out_Stream; at_row, at_column : Positive; text : String) is
  begin
    if text'Length >= 2048 then
      raise Constraint_Error;
    end if;
    --  5.70 Note
    case xl.xl_format is
      --  when BIFF8 =>  --  https://msdn.microsoft.com/en-us/library/dd945371(v=office.12).aspx
      --    Write_Biff(xl, 16#001C#,
      --      Intel_16(Unsigned_16(row-1)) &
      --      Intel_16(Unsigned_16(column-1)) &
      --      (0, 0) &  --  Show / hide options
      --      (0, 0) --  idObj - it begins to be tough there...
      --    );
      when others =>
        Write_Biff (xl, 16#001C#,
          Intel_16 (Unsigned_16 (at_row - 1)) &
          Intel_16 (Unsigned_16 (at_column - 1)) &
          To_buf_16_bit_length (text)
        );
    end case;
  end Write_Cell_Comment;

  procedure Write_Cell_Comment_at_Cursor (xl : Excel_Out_Stream; text : String) is
  begin
    Write_Cell_Comment (xl, Row (xl), Column (xl), text);
  end Write_Cell_Comment_at_Cursor;

  procedure Put_Line (xl : in out Excel_Out_Stream; num : Long_Float) is
  begin
    Put (xl, num);
    New_Line (xl);
  end Put_Line;

  procedure Put_Line (xl : in out Excel_Out_Stream; num : Integer) is
  begin
    Put (xl, num);
    New_Line (xl);
  end Put_Line;

  procedure Put_Line (xl : in out Excel_Out_Stream; str : String) is
  begin
    Put (xl, str);
    New_Line (xl);
  end Put_Line;

  procedure Put_Line (xl : in out Excel_Out_Stream; str : Ada.Strings.Unbounded.Unbounded_String) is
  begin
    Put_Line (xl, To_String (str));
  end Put_Line;

  procedure Put_Line (xl : in out Excel_Out_Stream; date : Ada.Calendar.Time) is
  begin
    Put (xl, date);
    New_Line (xl);
  end Put_Line;

  procedure New_Line (xl : in out Excel_Out_Stream; Spacing : Positive := 1) is
  begin
    Jump_to (xl, xl.curr_row + Spacing, 1);
  end New_Line;

  function Col (xl : in Excel_Out_Stream) return Positive is
  begin
    return xl.curr_col;
  end Col;

  function Column (xl : in Excel_Out_Stream) return Positive renames Col;

  function Line (xl : in Excel_Out_Stream) return Positive is
  begin
    return xl.curr_row;
  end Line;

  function Row (xl : in Excel_Out_Stream) return Positive renames Line;

  --  Relative / absolute jumps
  procedure Jump (xl : in out Excel_Out_Stream; rows, columns : Natural) is
  begin
    Jump_to (xl, xl.curr_row + rows, xl.curr_col + columns);
  end Jump;

  procedure Jump_to (xl : in out Excel_Out_Stream; to_row, to_column : Positive) is
  begin
    if to_row < xl.curr_row then -- trying to overwrite cells ?...
      raise Decreasing_row_index;
    end if;
    if to_row = xl.curr_row and then
      to_column < xl.curr_col
    then -- trying to overwrite cells on same row ?...
      raise Decreasing_column_index;
    end if;
    if to_row > 65536 then
      raise Row_out_of_range;
    elsif to_column > 256 then
      raise Column_out_of_range;
    end if;
    xl.curr_row := to_row;
    xl.curr_col := to_column;
  end Jump_to;

  procedure Next (xl : in out Excel_Out_Stream; columns : Natural := 1) is
  begin
    Jump (xl, rows => 0, columns => columns);
  end Next;

  procedure Next_Row (xl : in out Excel_Out_Stream; rows : Natural := 1) is
  begin
    Jump (xl, rows => rows, columns => 0);
  end Next_Row;

  procedure Use_Format (
    xl           : in out Excel_Out_Stream;
    format       : in     Format_Type
  )
  is
  begin
    if Integer (format) in XF_Range'Range then
      xl.xf_in_use := XF_Range (format);
    else
      raise Format_out_of_range;
      --  ^ Raised only if `format` was hacked using Unchecked_Conversion.
    end if;
  end Use_Format;

  procedure Use_Default_Format (xl : in out Excel_Out_Stream) is
  begin
    Use_Format (xl, xl.def_fmt);
  end Use_Default_Format;

  function Default_Font (xl : Excel_Out_Stream) return Font_Type is
  begin
    return xl.def_font;
  end Default_Font;

  function Default_Format (xl : Excel_Out_Stream) return Format_Type is
  begin
    return xl.def_fmt;
  end Default_Format;

  procedure Freeze_Panes (xl : in out Excel_Out_Stream; at_row, at_column : Positive) is
  begin
    xl.frz_panes := True;
    xl.freeze_row := at_row;
    xl.freeze_col := at_column;
  end Freeze_Panes;

  procedure Freeze_Panes_at_Cursor (xl : in out Excel_Out_Stream) is
  begin
    Freeze_Panes (xl, xl.curr_row, xl.curr_col);
  end Freeze_Panes_at_Cursor;

  procedure Freeze_Top_Row (xl : in out Excel_Out_Stream) is
  begin
    Freeze_Panes (xl, 2, 1);
  end Freeze_Top_Row;

  procedure Freeze_First_Column (xl : in out Excel_Out_Stream) is
  begin
    Freeze_Panes (xl, 1, 2);
  end Freeze_First_Column;

  procedure Zoom_Level (xl : in out Excel_Out_Stream; numerator, denominator : Positive) is
  begin
    xl.zoom_num := numerator;
    xl.zoom_den := denominator;
  end Zoom_Level;

  procedure Reset (
    xl           : in out Excel_Out_Stream'Class;
    excel_format :        Excel_Type;
    encoding     :        Encoding_Type
  )
  is
    dummy_xl_with_defaults : Excel_Out_Pre_Root_Type;
  begin
    --  Check if we are trying to re-use a half-finished object (ouch!):
    if xl.is_created and not xl.is_closed then
      raise Excel_stream_not_closed;
    end if;
    --  We will reset everything with defaults, except this:
    dummy_xl_with_defaults.xl_format := excel_format;
    dummy_xl_with_defaults.encoding  := encoding;
    --  Now we reset xl:
    Excel_Out_Pre_Root_Type (xl) := dummy_xl_with_defaults;
  end Reset;

  procedure Finish (xl : in out Excel_Out_Stream'Class) is

    procedure Write_Window1 is
    begin
      --  5.109 WINDOW1, p. 215
      case xl.xl_format is
        when BIFF2 | BIFF3 | BIFF4 =>  --  NB: more options in BIFF8
          Write_Biff (xl, 16#003D#,
            Intel_16 (120)   & -- Window x
            Intel_16 (120)   & -- Window y
            Intel_16 (21900) & -- Window w
            Intel_16 (13425) & -- Window h
            Intel_16 (0)       -- Hidden
          );
      end case;
    end Write_Window1;

    procedure Write_Window2 is
    begin
      --  5.110 WINDOW2
      case xl.xl_format is
        when BIFF2 =>
          Write_Biff (xl, 16#003E#,
            (0, -- Display formulas, not results
             1, -- Show grid lines
             1, -- Show sheet headers
             Boolean'Pos (xl.frz_panes),
             1  -- Show zero values as zeros, not empty cells
            )
             &
            Intel_16 (0) & -- First visible row
            Intel_16 (0) & -- First visible column
            (1, -- Use automatic grid line colour
             0, 0, 0, 0) -- Grid line RGB colour
          );
        when BIFF3 | BIFF4 =>  --  NB: more options in BIFF8
          Write_Biff (xl, 16#023E#,
            --  http://msdn.microsoft.com/en-us/library/dd947893(v=office.12).aspx
            Intel_16 (   -- Option flags:
              0 *   1 + -- Display formulas, not results
              1 *   2 + -- Show grid lines
              1 *   4 + -- Show sheet headers
              Boolean'Pos (xl.frz_panes)
                *   8 + -- Panes are frozen
              1 *  16 + -- Show zero values as zeros, not empty cells
              1 *  32 + -- Gridlines of the window drawn in the default window foreground color
              0 *  64 + -- Right-to-left mode
              1 * 128 + -- Show outlines (guts ?!)
              0 * 256   -- Frozen, not split
            ) &
            Intel_16 (0) & -- First visible row
            Intel_16 (0) & -- First visible column
            Intel_32 (0)   -- Grid line colour
          );
      end case;
    end Write_Window2;

    procedure Write_Pane is
      active_pane : Unsigned_8;
    begin
      if xl.freeze_col = 1 then
        if xl.freeze_row = 1 then
          active_pane := 3;
        else
          active_pane := 2;
        end if;
      else
        if xl.freeze_row = 1 then
          active_pane := 1;
        else
          active_pane := 0;
        end if;
      end if;
      --  5.75 PANE
      Write_Biff (xl, 16#0041#,
        Intel_16 (Unsigned_16 (xl.freeze_col) - 1) &
        Intel_16 (Unsigned_16 (xl.freeze_row) - 1) &
        Intel_16 (Unsigned_16 (xl.freeze_row) - 1) &
        Intel_16 (Unsigned_16 (xl.freeze_col) - 1) &
        (1 => active_pane)
      );
    end Write_Pane;

    col_bits : Byte_Buffer (1 .. 32) := (others => 0);
    byte_idx, bit_idx : Positive := 1;

  begin
    --  Calling Window1 and Window2 is not necessary for default settings, but without these calls,
    --  a Write_row_height call with a positive height results, on all MS Excel versions, in a
    --  completely blank row, including the header letters - clearly an Excel bug !
    Write_Window1;
    Write_Window2;
    --  5.92 SCL = Zoom, Magnification. Defined for BIFF4+ only, but works with BIFF2, BIFF3.
    Write_Biff (xl, 16#00A0#,
      Intel_16 (Unsigned_16 (xl.zoom_num)) &
      Intel_16 (Unsigned_16 (xl.zoom_den))
    );
    if xl.frz_panes and xl.xl_format > BIFF2 then
      --  Enabling PANE for BIFF2 causes a very strange behaviour on MS Excel 2002.
      Write_Pane;
    end if;
    --  5.93 SELECTION here !!
    if xl.xl_format >= BIFF4 then
      for i in 1 .. 256 loop
        col_bits (byte_idx) := col_bits (byte_idx) +
          Boolean'Pos (xl.std_col_width (i)) * (2**(bit_idx - 1));
        bit_idx := bit_idx + 1;
        if bit_idx = 9 then
          bit_idx := 1;
          byte_idx := byte_idx + 1;
        end if;
      end loop;
      --  5.51 GCW: Global Column Width - trying to get a correct display by LibreOffice
      --  Result: OK but useless on MS Excel, not working on LibreOffice :-(
      Write_Biff (xl, 16#00AB#, Intel_16 (32) & col_bits);
      --  if xl.defcolwdth > 0 then
      --    --  5.101 STANDARDWIDTH -- this confuses MS Excel...
      --    Write_Biff(xl, 16#0099#, Intel_16(Unsigned_16(xl.defcolwdth)));
      --  end if;
    end if;
    --  5.37 EOF: End of File:
    Write_Biff (xl, 16#000A#, empty_buffer);
    Set_Index (xl, xl.dimrecpos); -- Go back to overwrite the DIMENSION record with correct data
    Write_Dimensions (xl);
    xl.is_closed := True;
  end Finish;

  ----------------------
  -- Output to a file --
  ----------------------

  procedure Create (
    xl           : in out Excel_Out_File;
    file_name    :        String;
    excel_format :        Excel_Type    := Default_Excel_Type;
    encoding     :        Encoding_Type := Default_Encoding
  )
  is
    use Ada.Streams, Ada.Streams.Stream_IO;
  begin
    Reset (xl, excel_format, encoding);
    xl.xl_file := new Ada.Streams.Stream_IO.File_Type;
    Create (xl.xl_file.all, Out_File, file_name);
    xl.xl_stream := XL_Raw_Stream_Class (Stream (xl.xl_file.all));
    Write_Worksheet_header (xl);
  end Create;

  procedure Close (xl : in out Excel_Out_File) is
    procedure Dispose is new
      Ada.Unchecked_Deallocation (Ada.Streams.Stream_IO.File_Type, XL_File_Acc);
  begin
    Finish (xl);
    Ada.Streams.Stream_IO.Close (xl.xl_file.all);
    Dispose (xl.xl_file);
  end Close;

  --  Set the index on the file
  procedure Set_Index (xl : in out Excel_Out_File;
                       To : Ada.Streams.Stream_IO.Positive_Count)
  is
  begin
    Ada.Streams.Stream_IO.Set_Index (xl.xl_file.all, To);
  end Set_Index;

  --  Return the index of the file
  function Index (xl : Excel_Out_File) return Ada.Streams.Stream_IO.Count
  is
  begin
    return Ada.Streams.Stream_IO.Index (xl.xl_file.all);
  end Index;

  function Is_Open (xl : in Excel_Out_File) return Boolean is
  begin
    if xl.xl_file = null then
      return False;
    end if;
    return Ada.Streams.Stream_IO.Is_Open (xl.xl_file.all);
  end Is_Open;

  ------------------------
  -- Output to a string --
  ------------------------
  --  Code reused from Zip_Streams

  procedure Read
    (Stream : in out Unbounded_Stream;
     Item   :    out Ada.Streams.Stream_Element_Array;
     Last   :    out Ada.Streams.Stream_Element_Offset)
  is
    use Ada.Streams;
  begin
    --  Item is read from the stream. If (and only if) the stream is
    --  exhausted, Last will be < Item'Last. In that case, T'Read will
    --  raise an End_Error exception.
    --
    --  Cf: RM 13.13.1(8), RM 13.13.1(11), RM 13.13.2(37) and
    --  explanations by Tucker Taft
    --
    Last := Item'First - 1;
    --  if Item is empty, the following loop is skipped; if Stream.Loc
    --  is already indexing out of Stream.Unb, that value is also appropriate
    for i in Item'Range loop
      Item (i) := Character'Pos (Element (Stream.Unb, Stream.Loc));
      Stream.Loc := Stream.Loc + 1;
      Last := i;
    end loop;
  exception
    when Ada.Strings.Index_Error =>
      null; -- what could be read has been read; T'Read will raise End_Error
  end Read;

  procedure Write
    (Stream : in out Unbounded_Stream;
     Item   : Ada.Streams.Stream_Element_Array)
  is
  begin
    for I in Item'Range loop
      if Length (Stream.Unb) < Stream.Loc then
        Append (Stream.Unb, Character'Val (Item (I)));
      else
        Replace_Element (Stream.Unb, Stream.Loc, Character'Val (Item (I)));
      end if;
      Stream.Loc := Stream.Loc + 1;
    end loop;
  end Write;

  procedure Set_Index (S : access Unbounded_Stream; To : Positive) is
  begin
    if Length (S.Unb) < To then
      for I in Length (S.Unb) .. To loop
        Append (S.Unb, ASCII.NUL);
      end loop;
    end if;
    S.Loc := To;
  end Set_Index;

  function Index (S : access Unbounded_Stream) return Integer is
  begin
    return S.Loc;
  end Index;

  --- ***

  procedure Create (
    xl           : in out Excel_Out_String;
    excel_format :        Excel_Type    := Default_Excel_Type;
    encoding     :        Encoding_Type := Default_Encoding
  )
  is
  begin
    Reset (xl, excel_format, encoding);
    xl.xl_memory := new Unbounded_Stream;
    xl.xl_memory.Unb := Null_Unbounded_String;
    xl.xl_memory.Loc := 1;
    xl.xl_stream := XL_Raw_Stream_Class (xl.xl_memory);
    Write_Worksheet_header (xl);
  end Create;

  procedure Close (xl : in out Excel_Out_String) is
  begin
    Finish (xl);
  end Close;

  function Contents (xl : Excel_Out_String) return String is
  begin
    if not xl.is_closed then
      raise Excel_stream_not_closed;
    end if;
    return To_String (xl.xl_memory.Unb);
  end Contents;

  --  Set the index on the Excel string stream
  procedure Set_Index (xl : in out Excel_Out_String;
                       To : Ada.Streams.Stream_IO.Positive_Count)
  is
  begin
    Set_Index (xl.xl_memory, Integer (To));
  end Set_Index;

  --  Return the index of the Excel string stream
  function Index (xl : Excel_Out_String) return Ada.Streams.Stream_IO.Count
  is
  begin
    return Ada.Streams.Stream_IO.Count (Index (xl.xl_memory));
  end Index;

  function "&"(a, b : Font_Style) return Font_Style is
  begin
    return a or b; -- "or" is predefined for sets (=array of Boolean)
  end "&";

  function "&"(a, b : Cell_Border) return Cell_Border is
  begin
    return a or b; -- "or" is predefined for sets (=array of Boolean)
  end "&";

end Excel_Out;
