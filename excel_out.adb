-- Derived from ExcelOut @ http://www.modula2.org/projects/excelout.php
-- by Frank Schoonjans - thanks!
--
-- Translated with Mod2Pas and P2Ada
--
-- References to documentation are to http://sc.openoffice.org/excelfileformat.pdf
--
-- To do:
-- =====
--  - freeze pane (5.75 PANE)
--  - BIFF > 2 support
--  - ...

with Ada.Unchecked_Deallocation;
with Ada.Streams.Stream_IO;             use Ada.Streams.Stream_IO, Ada.Streams;
with Ada.Strings.Fixed;                 use Ada.Strings.Fixed;

with Interfaces;                        use Interfaces;

-- IEEE_754 from: Simple components for Ada by Dmitry A. Kazakov
-- http://www.dmitry-kazakov.de/ada/components.htm
with IEEE_754.Long_Floats;

package body Excel_Out is

  -- Very low level part which deals with transfering data endian-proof,
  -- and floats in the ieee format.

  type Byte_Buffer is array (Integer range <>) of Unsigned_8;

  function To_buf(s: String) return Byte_Buffer is
    b: Byte_Buffer(s'Range);
  begin
    if s'Length > 255 then -- length doesn't fit in a byte
      raise Constraint_Error;
    end if;
    for i in b'Range loop
      b(i):= Character'Pos(s(i));
    end loop;
    return Unsigned_8(s'Length) & b;
  end To_buf;
  -- Put numbers with correct endianess as bytes

  generic
    type Number is mod <>; -- range <> in Ada83 version (fake Interfaces)
    size: Positive;
  function Intel_x86_buffer( n: Number ) return Byte_Buffer;

  function Intel_x86_buffer( n: Number ) return Byte_Buffer is
    b: Byte_Buffer(1..size);
    m: Number:= n;
  begin
    for i in b'Range loop
      b(i):= Unsigned_8(m and 255);
      m:= m / 256;
    end loop;
    return b;
  end Intel_x86_buffer;

  function Intel_16 is new Intel_x86_buffer( Unsigned_16, 2 );

  -- Gives a byte sequence of an IEEE 64-bit number as if taken
  -- from an Intel machine (with the same endianess).
  --
  -- http://en.wikipedia.org/wiki/IEEE_754-1985#Double-precision_64_bit
  --

  function IEEE_Double_Intel(x: Long_Float) return Byte_buffer is
    subtype LF_bytes is Byte_buffer(1..8);
    d : LF_bytes;
    --
    use IEEE_754.Long_Floats;
    f64: constant Float_64:= To_IEEE(x);
  begin
    for i in d'Range loop
      d(i):= f64(9-i); -- Order is reversed
    end loop;
    -- Fully tested in Test_IEEE.adb
    return d;
  end IEEE_Double_Intel;

  ----------------
  -- Excel BIFF --
  ----------------

  -- The original code counts on a certain record packing & endianess.
  -- We do it without this assumption.

  procedure WriteBiff(
    xl     : Excel_Out_Stream'Class;
    biff_id: Unsigned_16;
    data   : Byte_buffer
  )
  is
  begin
    Byte_Buffer'Write(xl.xl_stream, Intel_16(biff_id));
    Byte_Buffer'Write(xl.xl_stream, Intel_16(Unsigned_16(data'Length)));
    Byte_Buffer'Write(xl.xl_stream, data);
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
    end case;
  end WriteBOF;

  -- 5.37 EOF: End of File
  procedure WriteEOF(xl : Excel_Out_Stream'Class) is
  begin
    WriteBiff(xl, 16#000A#, (1..0 => 0));
  end WriteEOF;

  procedure WriteFmtRecords (xl : Excel_Out_Stream'Class) is

    -- 5.49 FORMAT (number format)
    procedure WriteFmtStr (s : String) is
    begin
      case xl.format is
        when BIFF2 =>
          WriteBiff(xl, 16#001E#, To_buf(s));
      end case;
    end WriteFmtStr;

    fmt_count: constant:= 3;

  begin
    -- 5.12 BUILTINFMTCOUNT
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#001F#, Intel_16(fmt_count));
    end case;
    -- loop & case avoid omitting any choice
    for n in Number_format_type loop
      case n is
        when general    =>  WriteFmtStr("General");
        when decimal_0  =>  WriteFmtStr("0");
        when decimal_2  =>  WriteFmtStr("0.00");
        when decimal_0_thousands_separator =>
                            WriteFmtStr("#'##0");
        when decimal_2_thousands_separator =>     -- 'Comma' built-in style
                            WriteFmtStr("#'##0.00");
        when percent_0  =>  WriteFmtStr("0%");    -- 'Percent' built-in style
        when percent_2  =>  WriteFmtStr("0.00%");
        when percent_0_plus  =>
          WriteFmtStr("+0%;-0%;0%");
        when percent_2_plus  =>
          WriteFmtStr("+0.00%;-0.00%;0.00%");
        when scientific =>  WriteFmtStr("0.00E+00");
      end case;
    end loop;
    -- ^ Some formats in the original list caused problems, probably
    --   because of regional placeholder symbols
  end WriteFmtRecords;

  -- 5.35 DIMENSION
  procedure WriteDimensions(xl: Excel_Out_Stream'Class) is
  begin
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#0000#,
          Intel_16(0) & Intel_16(Unsigned_16(xl.maxrow + 1)) &
          Intel_16(0) & Intel_16(Unsigned_16(xl.maxcolumn + 1))
        );
    end case;
  end WriteDimensions;

  procedure Write_Worksheet_header(xl : in out Excel_Out_Stream'Class) is
  begin
    WriteBOF(xl);
    -- 5.17 CODEPAGE
    WriteBiff(xl, 16#0042#, Intel_16(16#8001#)); -- Windows CP-1252
    -- 5.14 CALCMODE
    WriteBiff(xl, 16#000D#, Intel_16(1)); --  1 = automatic
    -- 5.85 REFMODE
    WriteBiff(xl, 16#000F#, Intel_16(1)); --  1 = A1 mode
    -- 5.28 DATEMODE
    WriteBiff(xl, 16#0022#, Intel_16(1)); --  1 => 1904 Date system
    --
    WriteFmtRecords(xl);
    xl.dimrecpos:= Index(xl);
    WriteDimensions(xl);
    Define_font(xl,"Arial", 10, xl.def_font);
    Define_format(xl, xl.def_font, general, xl.def_fmt);
    xl.is_created:= True;
  end Write_Worksheet_header;

  -- *** Exported procedures **********************************************

  -- 5.115 XF - Extended Format
  procedure Define_format(
    xl           : in out Excel_Out_Stream;
    font         : in     Font_type;   -- given by Define_font
    number_format: in     Number_format_type;
    format       :    out Format_type;
    -- optional:
    horiz_align  : in     Horizontal_alignment:= general;
    border       : in     Cell_border:= no_border;
    shaded       : in     Boolean:= False
  )
  is
    border_bits, mask: Unsigned_8;
  begin
    case xl.format is
      when BIFF2 => -- 5.115.2 XF Record Contents
        border_bits:= 0;
        mask:= 8;
        for s in Cell_border_single loop
          if border(s) then
            border_bits:= border_bits + mask;
          end if;
          mask:= mask * 2;
        end loop;
        WriteBiff(xl, 16#0043#,
          (Unsigned_8(font),
           -- ^ Index to FONT record
           0,
           -- ^ Not used
           Number_format_type'Pos(number_format),
           -- ^ Number format and cell flags
           Horizontal_alignment'Pos(horiz_align) +
           border_bits +
           Boolean'Pos(shaded) * 128
           -- ^ Horizontal alignment, border style, and background
          )
        );
    end case;
    xl.xfs:= xl.xfs + 1;
    format:= Format_Type(xl.xfs);
  end Define_Format;

  y_scale: constant:= 20; -- scaling to obtain character point (pt) units

  -- 5.32 DEFAULTROWHEIGHT
  procedure Write_default_row_height (
        xl     : Excel_Out_Stream;
        height : Positive
  )
  is
  begin
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#0025#,
          Intel_16(Unsigned_16(height * y_scale))
        );
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

  -- 5.20 COLWIDTH (BIFF2 only)
  procedure Write_column_width (
        xl : Excel_Out_Stream;
        column : Positive;
        width  : Natural)
  is
  begin
    WriteBiff(xl, 16#0024#,
      Unsigned_8(column-1) & -- first
      Unsigned_8(column-1) & -- last
      Intel_16(Unsigned_16(width * 256))
    );
  end Write_column_width;

  -- 5.88 ROW
  -- The OpenOffice documentation tells nice stories about row blocks,
  -- but single ROW commands can also be put before in the data stream,
  -- where the column widths are set. Excel saves with blocks of ROW
  -- commands, most of them useless.

  procedure Write_row_height(
    xl : Excel_Out_Stream;
    row: Positive; height : Natural
  )
  is
  begin
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#0008#,
          Intel_16(Unsigned_16(row-1)) &
          Intel_16(0)   & -- col. min.
          Intel_16(256) & -- col. max. - we just take the full range...
          Intel_16(Unsigned_16(height * y_scale)) &
          (1..5=> 0)
        );
    end case;
  end Write_row_height;

  -- 5.45 FONT
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
    colcode: constant array(Color_type) of Unsigned_16:=
      (
         black     => 0,
         white     => 1,
         red       => 2,
         green     => 3,
         blue      => 4,
         yellow    => 5,
         magenta   => 6,
         cyan      => 7,
         automatic => 16#7FFF# -- system window text colour
      );
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
      xl.fonts:= 5; -- anomaly in all BIFF versions...
    end if;
    case xl.format is
      when BIFF2 =>
        WriteBiff(xl, 16#0031#,
          Intel_16(Unsigned_16(height * y_scale)) &
          Intel_16(style_bits) &
          To_buf(font_name)
        );
        if color /= automatic then
          -- 5.47 FONTCOLOR
          WriteBiff(xl, 16#0045#, Intel_16(colcode(color)));
        end if;
    end case;
    font:= Font_Type(xl.fonts);
  end Define_font;

  procedure StoreMaxRC(xl: in out Excel_Out_Stream; r, c: Integer) is
  begin
    if not xl.is_created then
      raise Excel_Stream_Not_Created;
    end if;
    if r > xl.maxrow then
      xl.maxrow := r;
    end if;
    if c > xl.maxcolumn then
      xl.maxcolumn := c;
    end if;
  end StoreMaxRC;

  -- 2.5.13 Cell Attributes (BIFF2 only)
  function Cell_attributes(xl: Excel_Out_Stream) return Byte_buffer is
  begin
    return
      (Unsigned_8(xl.xf_in_use),
       0,
       0
      );
  end Cell_attributes;

  function Almost_zero(x: Long_Float) return Boolean is
  begin
    return abs x <= Long_Float'Model_Small;
  end Almost_zero;

  -- 5.71 NUMBER
  procedure Write (
        xl     : in out Excel_Out_Stream;
        r,
        c      : Positive;
        num    : Long_Float
  )
  is
  begin
    if xl.format = BIFF2 and then
       num >= 0.0 and then
       num <= 65535.0 and then
       Almost_zero(num - Long_Float'Floor(num))
    then
      Write(xl,r,c,Integer(Long_Float'Floor(num)));
    else
      StoreMaxRC(xl, r-1, c-1);
      Jump_to(xl, r,c); -- Store and check current position
      case xl.format is
        when BIFF2 =>
          WriteBiff(xl, 16#0003#,
            Intel_16(Unsigned_16(r-1)) &
            Intel_16(Unsigned_16(c-1)) &
            Cell_attributes(xl) &
            IEEE_Double_Intel(num)
          );
      end case;
      Jump_to(xl, r,c+1); -- Store and check new position
    end if;
  end Write;

  procedure Write (
        xl : in out Excel_Out_Stream;
        r,
        c      : Positive;
        num    : Integer)
  is
  begin
    if xl.format = BIFF2 and then
       num in 0..2**16-1
    then -- We use a small storage for integers
      Jump_to(xl, r,c); -- Store and check current position
      StoreMaxRC(xl, r-1, c-1);
      -- 5.60 INTEGER
      WriteBiff(xl, 16#0002#,
        Intel_16(Unsigned_16(r-1)) &
        Intel_16(Unsigned_16(c-1)) &
        Cell_attributes(xl) &
        Intel_16(Unsigned_16(num))
      );
      Jump_to(xl, r,c+1); -- Store and check new position
    else -- We need to us a floating-point
      Write(xl, r, c, Long_Float(num));
    end if;
  end Write;

  procedure Write (
        xl : in out Excel_Out_Stream;
        r,
        c      : Positive;
        str    : String)
  is
  begin
    Jump_to(xl, r,c); -- Store and check current position
    StoreMaxRC(xl, r-1, c-1);
    if str'Length = 0 then
      return;
    end if;
    case xl.format is
      when BIFF2 =>
        -- 5.63 LABEL
        WriteBiff(xl, 16#0004#,
          Intel_16(Unsigned_16(r-1)) &
          Intel_16(Unsigned_16(c-1)) &
          Cell_attributes(xl) &
          To_buf(str)
        );
    end case;
    Jump_to(xl, r,c+1); -- Store and check new position
  end Write;

  procedure Write(xl: in out Excel_Out_Stream; r,c : Positive; str : Unbounded_String)
  is
  begin
    Write(xl, r,c, To_String(str));
  end;

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
        s: String(1..50 + 0*width);
        -- 0*width is just to skip a warning of width being unused
        package IIO is new Ada.Text_IO.Integer_IO(Integer);
      begin
        IIO.Put(s, num, base => base);
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

  procedure New_Line(xl: in out Excel_Out_Stream) is
  begin
    Jump_to(xl, xl.curr_row + 1, 1);
  end New_Line;

  -- Relative / absolute jumps
  procedure Jump(xl: in out Excel_Out_Stream; rows, columns: Natural) is
  begin
    Jump_to(xl, xl.curr_row + rows, xl.curr_col + columns);
  end;

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
  end;

  procedure Use_format(
    xl           : in out Excel_Out_Stream;
    format       : in     Format_type
  )
  is
  begin
    xl.xf_in_use:= XF_Range(format);
  end Use_Format;

  procedure Use_default_format(xl: in out Excel_Out_Stream) is
  begin
    Use_format(xl, xl.def_fmt);
  end Use_default_format;

  function Default_font(xl: Excel_Out_Stream) return Font_type is
  begin
    return xl.def_font;
  end;

  function Default_format(xl: Excel_Out_Stream) return Format_type is
  begin
    return xl.def_fmt;
  end;

  procedure Reset(
    xl        : in out Excel_Out_Stream'Class;
    format    :        Excel_type:= Default_Excel_type
  )
  is
    dummy_xl_with_defaults: Excel_Out_Pre_Root_Type;
  begin
    -- Check if we are trying to re-use a half-finished object (ouch!):
    if xl.is_created and not xl.is_closed then
      raise Excel_Stream_Not_Closed;
    end if;
    dummy_xl_with_defaults.format:= format;
    Excel_Out_Pre_Root_Type(xl):= dummy_xl_with_defaults;
  end Reset;

  procedure Finish(xl : in out Excel_Out_Stream'Class) is
  begin
    WriteEOF(xl);
    Set_Index(xl, xl.dimrecpos);
    WriteDimensions(xl);
    xl.is_closed:= True;
  end Finish;

  ----------------------
  -- Output to a file --
  ----------------------

  procedure Create(
    xl        : in out Excel_Out_File;
    file_name :        String;
    format    :        Excel_type:= Default_Excel_type
  )
  is
  begin
    Reset(xl, format);
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
                       to: Ada.Streams.Stream_IO.Positive_Count)
  is
  begin
    Ada.Streams.Stream_IO.Set_Index(xl.xl_file.all, To);
  end;

  -- Return the index of the file
  function Index (xl: Excel_Out_File) return Ada.Streams.Stream_IO.Count
  is
  begin
    return Ada.Streams.Stream_IO.Index(xl.xl_file.all);
  end;

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
    xl        : in out Excel_Out_String;
    format    :        Excel_type:= Default_Excel_type
  )
  is
  begin
    Reset(xl, format);
    xl.xl_memory:= new Unbounded_Stream;
    xl.xl_memory.unb:= Null_Unbounded_String;
    xl.xl_memory.loc:= 1;
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
      raise Excel_Stream_Not_Closed;
    end if;
    return To_String(xl.xl_memory.unb);
  end Contents;

  -- Set the index on the Excel string stream
  procedure Set_Index (xl: in out Excel_Out_String;
                       to: Ada.Streams.Stream_IO.Positive_Count)
  is
  begin
    Set_Index(xl.xl_memory, Integer(to));
  end;

  -- Return the index of the Excel string stream
  function Index (xl: Excel_Out_String) return Ada.Streams.Stream_IO.Count
  is
  begin
    return Ada.Streams.Stream_IO.Count(Index(xl.xl_memory));
  end;

end Excel_Out;
