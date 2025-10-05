-------------------------------------------------------------------------------------
--
--  EXCEL_OUT - A low level package for writing Microsoft Excel (*) files
--
--  Pure Ada 95 code, 100% portable: OS-, CPU- and compiler- independent.
--
--  Version / date / download info: see the version, reference, web strings
--   defined at the end of the public part of this package.

--  Legal licensing note:

--  Copyright (c) 2009 .. 2025 Gautier de Montmollin

--  Permission is hereby granted, free of charge, to any person obtaining a copy
--  of this software and associated documentation files (the "Software"), to deal
--  in the Software without restriction, including without limitation the rights
--  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
--  copies of the Software, and to permit persons to whom the Software is
--  furnished to do so, subject to the following conditions:

--  The above copyright notice and this permission notice shall be included in
--  all copies or substantial portions of the Software.

--  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
--  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
--  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
--  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
--  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
--  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
--  THE SOFTWARE.

--  NB: this is the MIT License, as found 12-Sep-2007 on the site
--  http://www.opensource.org/licenses/mit-license.php

--  (*) All Trademarks mentioned are properties of their respective owners.
-------------------------------------------------------------------------------------
--
--  Follow these steps to create an Excel spreadsheet stream:
--
--  1. Create
--
--  2. Optional settings, before any data output:
--     | Write_default_column_width
--     | Write_column_width for specific columns
--     | Write_default_row_height
--     | Write_row_height for specific rows
--     | Define_Font, then Define_Format
--
--  3. | Write(xl, row, column, data): row by row, column by column
--     | Put(xl, data)               : same, but column is auto-incremented
--     | New_Line(xl),...            : other "Text_IO"-like (full list below)
--     | Use_format, influences the format of data written next
--
--  4. Close
--
--  5. (Excel_Out_String only) function Contents returns the full .xls
--------------------------------------------------------------------------

with Ada.Calendar,
     Ada.Streams.Stream_IO,
     Ada.Strings.Unbounded,
     Ada.Text_IO;

with Interfaces;

package Excel_Out is

  -----------------------------------------------------------------
  -- The abstract Excel output stream root type.                 --
  -- From this package, you can use the following derived types: --
  --    * Excel_Out_File    : output in a file                   --
  --    * Excel_Out_String  : output in a string                 --
  -- Of course you can define your own derived types.            --
  -----------------------------------------------------------------

  type Excel_Out_Stream is abstract tagged private;

  ----------------------------------
  -- (2) Before any cell content: --
  ----------------------------------

  --  * Page layout for printing
  procedure Header (xl : Excel_Out_Stream; page_header_string : String);
  procedure Footer (xl : Excel_Out_Stream; page_footer_string : String);
  --
  procedure Left_Margin (xl : Excel_Out_Stream; inches : Long_Float);
  procedure Right_Margin (xl : Excel_Out_Stream; inches : Long_Float);
  procedure Top_Margin (xl : Excel_Out_Stream; inches : Long_Float);
  procedure Bottom_Margin (xl : Excel_Out_Stream; inches : Long_Float);
  procedure Margins
    (xl : Excel_Out_Stream;
     left_inches,
     right_inches,
     top_inches,
     bottom_inches : Long_Float);
  --
  procedure Print_Row_Column_Headers (xl : Excel_Out_Stream);
  procedure Print_Gridlines (xl : Excel_Out_Stream);
  --
  type Orientation_Choice is (landscape, portrait);
  type Scale_or_Fit_Choice is (scale, fit);
  procedure Page_Setup
    (xl                      : Excel_Out_Stream;
     scaling_percents        : Positive := 100;
     fit_width_with_n_pages  : Natural  := 1;  --  0: as many as possible
     fit_height_with_n_pages : Natural  := 1;  --  0: as many as possible
     orientation             : Orientation_Choice  := portrait;
     scale_or_fit            : Scale_or_Fit_Choice := scale);

  --  * The column width unit is as it appears in Excel when you resize a column.
  --      It is the width of a '0' in a standard font.
  procedure Write_Default_Column_Width (xl : in out Excel_Out_Stream; width : Positive);
  procedure Write_Column_Width (xl : in out Excel_Out_Stream; column : Positive; width : Natural);
  procedure Write_Column_Width
    (xl            : in out Excel_Out_Stream;
     first_column,
     last_column   : Positive;
     width         : Natural);

  --  * The row height unit is in font points, as appearing when you
  --      resize a row in Excel. A zero height means the row is hidden.
  procedure Write_Default_Row_Height (xl : Excel_Out_Stream; height : Positive);
  procedure Write_Row_Height (xl : Excel_Out_Stream; row : Positive; height : Natural);

  ----------------------
  -- Formatting cells --
  ----------------------
  --  A cell format is, as you can see in the format dialog
  --  in Excel, a combination of:
  --    - a number format
  --    - a set of alignements
  --    - a font
  --    - and other optional things to come here...
  --  Formats are user-defined except one which is predefined: Default_format

  type Format_Type is private;

  function Default_Format (xl : Excel_Out_Stream) return Format_Type;
  --  What you get when creating a new sheet in Excel: Default_font,...

  --  * Number format
  type Number_Format_Type is private;

  --  Built-in number formats
  general                       : constant Number_Format_Type;
  decimal_0                     : constant Number_Format_Type;
  decimal_2                     : constant Number_Format_Type;
  decimal_0_thousands_separator : constant Number_Format_Type;  --  1'234'000
  decimal_2_thousands_separator : constant Number_Format_Type;  --  1'234'000.00
  percent_0                     : constant Number_Format_Type;  --   3%, 0%, -4%
  percent_2                     : constant Number_Format_Type;
  percent_0_plus                : constant Number_Format_Type;  --  +3%, 0%, -4%
  percent_2_plus                : constant Number_Format_Type;
  scientific                    : constant Number_Format_Type;
  dd_mm_yyyy                    : constant Number_Format_Type;
  dd_mm_yyyy_hh_mm              : constant Number_Format_Type;
  hh_mm                         : constant Number_Format_Type;
  hh_mm_ss                      : constant Number_Format_Type;

  procedure Define_Number_Format
    (xl            : in out Excel_Out_Stream;
     format        :    out Number_Format_Type;
     format_string : in     String);

  --  * Fonts are user-defined, one is predefined: Default_font
  type Font_Type is private;

  function Default_Font (xl : Excel_Out_Stream) return Font_Type;
  --  Arial 10, regular, "automatic" color

  type Color_Type is
    (automatic,
     black, white, red, green, blue, yellow, magenta, cyan,
     dark_red, dark_green, dark_blue, olive, purple, teal, silver, grey);

  type Font_Style is private;

  --  For combining font styles (e.g.: bold & underlined):
  function "&" (a, b : Font_Style) return Font_Style;

  regular     : constant Font_Style;
  italic      : constant Font_Style;
  bold        : constant Font_Style;
  bold_italic : constant Font_Style;
  underlined  : constant Font_Style;
  struck_out  : constant Font_Style;
  shadowed    : constant Font_Style;
  condensed   : constant Font_Style;
  extended    : constant Font_Style;

  procedure Define_Font
    (xl           : in out Excel_Out_Stream;
     font_name    :        String;
     height       :        Positive;
     font         :    out Font_Type;
     --  Optional:
     style        :        Font_Style := regular;
     color        :        Color_Type := automatic);

  type Horizontal_Alignment is
    (general_alignment, to_left, centred, to_right, filled,
     justified, centred_across_selection, -- (BIFF4-BIFF8)
     distributed);  --  (BIFF8, Excel 10.0 ("XP") and later only)

  type Vertical_Alignment is
    (top_alignment, centred, bottom_alignment,
     justified,     --  (BIFF5-BIFF8)
     distributed);  --  (BIFF8, Excel 10.0 ("XP") and later only)

  type Text_Orientation is
    (normal,
     stacked,       --  vertical, top to bottom
     rotated_90,    --  vertical, rotated 90 degrees counterclockwise
     rotated_270);  --  vertical, rotated 90 degrees clockwise

  type Cell_Border is private;

  --  Operator for combining borders (e.g.: left & top):
  function "&" (a, b : Cell_Border) return Cell_Border;

  no_border : constant Cell_Border;
  left      : constant Cell_Border;
  right     : constant Cell_Border;
  top       : constant Cell_Border;
  bottom    : constant Cell_Border;
  box       : constant Cell_Border;

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
     text_orient      : in     Text_Orientation     := normal);

  ------------------------
  -- (3) Cell contents: --
  ------------------------

  --  Notes:
  --    - You need to write things with ascending row indexes, and with ascending
  --        column indexes within a row. Otherwise Excel issues a protest.
  --    - For strings starting with '=', Excel_Out attempts to turn the contents
  --        into a formula, exactly as if you were typing the cell content.

  procedure Write (xl : in out Excel_Out_Stream; r, c : Positive; num  : Long_Float);
  procedure Write (xl : in out Excel_Out_Stream; r, c : Positive; num  : Integer);
  procedure Write (xl : in out Excel_Out_Stream; r, c : Positive; str  : String);
  procedure Write (xl : in out Excel_Out_Stream; r, c : Positive; str  : Ada.Strings.Unbounded.Unbounded_String);
  procedure Write (xl : in out Excel_Out_Stream; r, c : Positive; date : Ada.Calendar.Time);

  --  "Ada.Text_IO"-like variants for output.
  --  No need to specify row & column each time.
  --  Write 'Put(x, content)' where x is an Excel_Out_Stream just
  --  as if x was a File_Type, and vice-versa.
  --
  procedure Put (xl : in out Excel_Out_Stream; num : Long_Float);
  procedure Put (xl    : in out Excel_Out_Stream;
                 num   : in Integer;
                 width : in Ada.Text_IO.Field := 0;  --  ignored
                 base  : in Ada.Text_IO.Number_Base := 10);
  procedure Put (xl : in out Excel_Out_Stream; str  : String);
  procedure Put (xl : in out Excel_Out_Stream; str  : Ada.Strings.Unbounded.Unbounded_String);
  procedure Put (xl : in out Excel_Out_Stream; date : Ada.Calendar.Time);
  --
  procedure Put_Line (xl : in out Excel_Out_Stream; num  : Long_Float);
  procedure Put_Line (xl : in out Excel_Out_Stream; num  : Integer);
  procedure Put_Line (xl : in out Excel_Out_Stream; str  : String);
  procedure Put_Line (xl : in out Excel_Out_Stream; str  : Ada.Strings.Unbounded.Unbounded_String);
  procedure Put_Line (xl : in out Excel_Out_Stream; date : Ada.Calendar.Time);
  --
  procedure New_Line (xl : in out Excel_Out_Stream; Spacing : Positive := 1);

  --  Get current column and row. The next Put will put contents in that cell.
  --
  function Col (xl : in Excel_Out_Stream) return Positive;     --  Text_IO naming
  function Column (xl : in Excel_Out_Stream) return Positive;  --  Excel naming
  function Line (xl : in Excel_Out_Stream) return Positive;    --  Text_IO naming
  function Row (xl : in Excel_Out_Stream) return Positive;     --  Excel naming

  --  Relative / absolute jumps
  procedure Jump (xl : in out Excel_Out_Stream; rows, columns : Natural);
  procedure Jump_to (xl : in out Excel_Out_Stream; to_row, to_column : Positive);
  procedure Next (xl : in out Excel_Out_Stream; columns : Natural := 1);   --  Jump 0 or more cells right
  procedure Next_Row (xl : in out Excel_Out_Stream; rows : Natural := 1);  --  Jump 0 or more cells down
  --
  --  Merge a certain amount of cells with the last one,
  --  right to that cell, on the same row.
  procedure Merge (xl : in out Excel_Out_Stream; cells : Positive);

  procedure Write_Cell_Comment (xl : Excel_Out_Stream; at_row, at_column : Positive; text : String);
  procedure Write_Cell_Comment_at_Cursor (xl : Excel_Out_Stream; text : String);

  --  Cells written after Use_Format will be using the given format,
  --  defined by Define_Format.
  procedure Use_Format
    (xl           : in out Excel_Out_Stream;
     format       : in     Format_Type);
  procedure Use_Default_Format (xl : in out Excel_Out_Stream);

  --  The Freeze Pane methods can be called anytime before Close
  procedure Freeze_Panes (xl : in out Excel_Out_Stream; at_row, at_column : Positive);
  procedure Freeze_Panes_at_Cursor (xl : in out Excel_Out_Stream);
  procedure Freeze_Top_Row (xl : in out Excel_Out_Stream);
  procedure Freeze_First_Column (xl : in out Excel_Out_Stream);

  --  Zoom level. Example: for 85%, call with parameters 85, 100.
  procedure Zoom_Level (xl : in out Excel_Out_Stream; numerator, denominator : Positive);

  --  Set_Index and Index are not directly useful for Excel_Out users.
  --  They are private indeed, but they must be visible (RM 3.9.3(10)).

  --  Set the index on the stream
  procedure Set_Index (xl : in out Excel_Out_Stream;
                       to : Ada.Streams.Stream_IO.Positive_Count)
  is abstract;

  --  Return the index of the stream
  function Index (xl : Excel_Out_Stream) return Ada.Streams.Stream_IO.Count
  is abstract;

  Excel_stream_not_created,
  Excel_stream_not_closed,
  Decreasing_row_index,
  Decreasing_column_index,
  Row_out_of_range,
  Column_out_of_range,
  Format_out_of_range,
  Font_out_of_range,
  Number_format_out_of_range : exception;

  type Excel_Type is
    (BIFF2,    --  Excel 2.1, 2,2
     BIFF3,    --  Excel 3.0
     BIFF4);   --  Excel 4.0
     --  BIFF5,    --  Excel 5.0 to 7.0
     --  BIFF8);   --  Excel 8.0 (97) to 11.0 (2003) - UTF-16 support

  Default_Excel_Type : constant Excel_Type := BIFF4;

  --  Assumed encoding for types Character and String (8-bit):

  type Encoding_Type is
    (Windows_CP_874,  --  Thai
     Windows_CP_932,  --  Japanese Shift-JIS
     Windows_CP_936,  --  Chinese Simplified GBK
     Windows_CP_949,  --  Korean (Wansung)
     Windows_CP_950,  --  Chinese Traditional BIG5
     Windows_CP_1250, --  Latin II (Central European)
     Windows_CP_1251, --  Cyrillic
     Windows_CP_1252, --  Latin I, superset of ISO 8859-1
     Windows_CP_1253, --  Greek
     Windows_CP_1254, --  Turkish
     Windows_CP_1255, --  Hebrew
     Windows_CP_1256, --  Arabic
     Windows_CP_1257, --  Baltic
     Windows_CP_1258, --  Vietnamese
     Windows_CP_1361, --  Korean (Johab)
     Apple_Roman);

  Default_Encoding : constant Encoding_Type := Windows_CP_1252;

  ---------------------------------------------------
  --  Here come two derived concrete stream types  --
  --  that are pre-defined in this package.        --
  ---------------------------------------------------

  -------------------------------
  --  Stream output to a file  --
  -------------------------------

  type Excel_Out_File is new Excel_Out_Stream with private;

  procedure Create
    (xl           : in out Excel_Out_File;
     file_name    :        String;
     excel_format :        Excel_Type    := Default_Excel_Type;
     encoding     :        Encoding_Type := Default_Encoding);

  procedure Close (xl : in out Excel_Out_File);

  function Is_Open (xl : in Excel_Out_File) return Boolean;

  ---------------------------------
  --  Stream output to a string  --
  ---------------------------------

  --  The output string is (mis)used as a byte buffer,
  --  to be compressed, packaged, transmitted, ...

  type Excel_Out_String is new Excel_Out_Stream with private;

  procedure Create
    (xl           : in out Excel_Out_String;
     excel_format :        Excel_Type    := Default_Excel_Type;
     encoding     :        Encoding_Type := Default_Encoding);

  procedure Close (xl : in out Excel_Out_String);

  function Contents (xl : Excel_Out_String) return String;

  --------------------------
  --  A little goodie...  --
  --------------------------

  --  Like x'Image, but without leading space.
  --
  function Img (x : Integer) return String;

  ----------------------------------------------------------------
  --  Information about this package - e.g. for an "about" box  --
  ----------------------------------------------------------------

  version   : constant String := "19";
  reference : constant String := "05-Oct-2025";
  --  Hopefully the latest version is at one of those URLs:
  web       : constant String := "http://excel-writer.sf.net/";
  web2 : constant String := "https://sourceforge.net/projects/excel-writer/";
  web3 : constant String := "https://github.com/zertovitch/excel-writer";
  web4 : constant String := "https://alire.ada.dev/crates/excel_writer";

private

  ----------------------------------------
  -- Raw Streams, with 'Read and 'Write --
  ----------------------------------------

  type XL_Raw_Stream_Class is access all Ada.Streams.Root_Stream_Type'Class;

  type Font_Type is new Natural;
  type Number_Format_Type is new Natural;
  type Format_Type is new Natural;

  subtype XF_Range is Integer range 0 .. 62;
  --  ^ After 62 we would need to use an IXFE (5.62)

  --  Theoretically, we would not need to memorize the XF informations
  --  and just give the XF identifier given with Format_type, but some
  --  versions of Excel with some locales mix up the font and numerical format
  --  when giving 0 for the cell attributes (see Cell_attributes, 2.5.13)
  --   Added Mar-2011.

  type XF_Info is record
    font : Font_Type;
    numb : Number_Format_Type;
  end record;

  type XF_Definition is array (XF_Range) of XF_Info;

  --  Built-in number formats
  general          : constant Number_Format_Type := 0;
  decimal_0        : constant Number_Format_Type := 1;
  decimal_2        : constant Number_Format_Type := 2;
  decimal_0_thousands_separator : constant Number_Format_Type := 3;  -- 1'234'000
  decimal_2_thousands_separator : constant Number_Format_Type := 4;  -- 1'234'000.00
  no_currency_0       : constant Number_Format_Type := 5;
  no_currency_red_0   : constant Number_Format_Type := 6;
  no_currency_2       : constant Number_Format_Type := 7;
  no_currency_red_2   : constant Number_Format_Type := 8;
  currency_0       : constant Number_Format_Type :=  9; -- 5 in BIFF2, BIFF3 (sob!)
  currency_red_0   : constant Number_Format_Type := 10; -- 6 in BIFF2, BIFF3...
  currency_2       : constant Number_Format_Type := 11;
  currency_red_2   : constant Number_Format_Type := 12;
  percent_0        : constant Number_Format_Type := 13;  --  3%, 0%, -4%
  percent_2        : constant Number_Format_Type := 14;
  scientific       : constant Number_Format_Type := 15;
  fraction_1       : constant Number_Format_Type := 16;
  fraction_2       : constant Number_Format_Type := 17;
  dd_mm_yyyy       : constant Number_Format_Type := 18; -- 14 in BIFF3, 12 in BIFF2 (re-sob!)
  dd_mmm_yy        : constant Number_Format_Type := 19; -- 15 in BIFF3, 13 in BIFF2...
  dd_mmm           : constant Number_Format_Type := 20;
  mmm_yy           : constant Number_Format_Type := 21;
  h_mm_AM_PM       : constant Number_Format_Type := 22;
  h_mm_ss_AM_PM    : constant Number_Format_Type := 23;
  hh_mm            : constant Number_Format_Type := 24;
  hh_mm_ss         : constant Number_Format_Type := 25;
  dd_mm_yyyy_hh_mm : constant Number_Format_Type := 26;
  --  End of Excel built-in formats
  last_built_in : constant Number_Format_Type := dd_mm_yyyy_hh_mm;

  percent_0_plus   : constant Number_Format_Type := 27;  --  +3%, 0%, -4%
  percent_2_plus   : constant Number_Format_Type := 28;
  date_iso         : constant Number_Format_Type := 29;  --  ISO 8601 format: 2014-03-16
  date_h_m_iso     : constant Number_Format_Type := 30;  --  date, hour, minutes
  date_h_m_s_iso   : constant Number_Format_Type := 31;  --  date, hour, minutes, seconds
  --  End of our custom formats
  last_custom_number_format : constant Number_Format_Type := date_h_m_s_iso;

  type Col_Width_Set is array (1 .. 256) of Boolean;

  --  We have a concrete type as hidden ancestor of the Excel_Out_Stream root
  --  type. A variable of that type is initialized with default values and
  --  can help re-initialize a Excel_Out_Stream when re-used several times.
  --  See the Reset procedure in body.
  --  The abstract Excel_Out_Stream could have default values, but using a
  --  variable of this type to reset values is not Ada compliant (LRM:3.9.3(8))
  --
  type Excel_Out_Pre_Root_Type is tagged record
    xl_stream  : XL_Raw_Stream_Class;
    xl_format  : Excel_Type    := Default_Excel_Type;
    encoding   : Encoding_Type := Default_Encoding;
    dimrecpos  : Ada.Streams.Stream_IO.Positive_Count;
    maxcolumn  : Positive :=  1;
    maxrow     : Positive :=  1;
    fonts      : Integer  := -1;  --  [-1..max_font]
    xfs        : Integer  := -1;  --  [-1..XF_Range'Last]
    xf_in_use  : XF_Range :=  0;
    xf_def     : XF_Definition;
    number_fmt : Number_Format_Type := last_custom_number_format;
    def_font   : Font_Type;
    def_fmt    : Format_Type;  --  Default format; used for "Normal" style
    cma_fmt    : Format_Type;  --  Format used for defining "Comma" style
    ccy_fmt    : Format_Type;  --  Format used for defining "Currency" style
    pct_fmt    : Format_Type;  --  Format used for defining "Percent" style
    is_created : Boolean := False;
    is_closed  : Boolean := False;
    curr_row   : Positive := 1;
    curr_col   : Positive := 1;
    frz_panes  : Boolean := False;
    freeze_row : Positive;
    freeze_col : Positive;
    zoom_num   : Positive := 100;
    zoom_den   : Positive := 100;
    defcolwdth : Natural := 0;  --  0 = not set; 1/256 of the width of the zero character
    std_col_width : Col_Width_Set := (others => True);
  end record;

  type Excel_Out_Stream is abstract new Excel_Out_Pre_Root_Type with null record;

  type Font_Style_Single is
    (bold_single,
     italic_single,
     underlined_single,
     struck_out_single,
     outlined_single,
     shadowed_single,
     condensed_single,
     extended_single);

  type Font_Style is array (Font_Style_Single) of Boolean;  --  This type is a set.

  regular     : constant Font_Style := (others => False);
  italic      : constant Font_Style := (italic_single => True, others => False);
  bold        : constant Font_Style := (bold_single => True, others => False);
  bold_italic : constant Font_Style := bold or italic;
  underlined  : constant Font_Style := (underlined_single => True, others => False);
  struck_out  : constant Font_Style := (struck_out_single => True, others => False);
  shadowed    : constant Font_Style := (shadowed_single => True, others => False);
  condensed   : constant Font_Style := (condensed_single => True, others => False);
  extended    : constant Font_Style := (extended_single => True, others => False);

  type Cell_Border_Single is
    (left_single,
     right_single,
     top_single,
     bottom_single);

  type Cell_Border is array (Cell_Border_Single) of Boolean;  --  This type is a set.

  no_border : constant Cell_Border := (others => False);
  left      : constant Cell_Border := (left_single => True, others => False);
  right     : constant Cell_Border := (right_single => True, others => False);
  top       : constant Cell_Border := (top_single => True, others => False);
  bottom    : constant Cell_Border := (bottom_single => True, others => False);
  box       : constant Cell_Border := (others => True);

  ----------------------
  -- Output to a file --
  ----------------------

  type XL_File_Acc is
    access Ada.Streams.Stream_IO.File_Type;

  type Excel_Out_File is new Excel_Out_Stream with record
    xl_file   : XL_File_Acc := null;  --  access to the "physical" Excel file
  end record;

  --  Set the index on the file
  procedure Set_Index (xl : in out Excel_Out_File;
                       To : Ada.Streams.Stream_IO.Positive_Count);

  --  Return the index of the file
  function Index (xl : Excel_Out_File) return Ada.Streams.Stream_IO.Count;

  ------------------------
  -- Output to a string --
  ------------------------
  --  Code reused from Zip_Streams

  --- *** We define here a complete in-memory stream:
  type Unbounded_Stream is new Ada.Streams.Root_Stream_Type with
    record
      Unb : Ada.Strings.Unbounded.Unbounded_String;
      Loc : Integer := 1;
    end record;

  --  Read data from the stream.
  procedure Read
    (Stream : in out Unbounded_Stream;
     Item   : out Ada.Streams.Stream_Element_Array;
     Last   : out Ada.Streams.Stream_Element_Offset);

  --  Write data to the stream, starting from the current index.
  --  Data will be overwritten from index is already available.
  procedure Write
    (Stream : in out Unbounded_Stream;
     Item   : Ada.Streams.Stream_Element_Array);

  --  Set the index on the stream
  procedure Set_Index (S : access Unbounded_Stream; To : Positive);

  --  returns the index of the stream
  function Index (S : access Unbounded_Stream) return Integer;

  --- ***

  type Unbounded_Stream_Acc is access Unbounded_Stream;

  type Excel_Out_String is new Excel_Out_Stream with record
    xl_memory : Unbounded_Stream_Acc;
  end record;

  --  Set the index on the Excel string stream
  procedure Set_Index (xl : in out Excel_Out_String;
                       To : Ada.Streams.Stream_IO.Positive_Count);

  --  Return the index of the Excel string stream
  function Index (xl : Excel_Out_String) return Ada.Streams.Stream_IO.Count;

  --  Very low level part which deals with transferring data byte-wise.

  type Byte_Buffer is array (Integer range <>) of Interfaces.Unsigned_8;
  empty_buffer : constant Byte_Buffer := (1 .. 0 => 0);

  --  Put numbers with correct endianess as bytes:
  generic
    type Number is mod <>;
    size : Positive;
  function Intel_x86_buffer (n : Number) return Byte_Buffer;
  pragma Inline (Intel_x86_buffer);

  function Intel_16 (n : Interfaces.Unsigned_16) return Byte_Buffer;
  pragma Inline (Intel_16);

  function Intel_32 (n : Interfaces.Unsigned_32) return Byte_Buffer;
  pragma Inline (Intel_32);

  function IEEE_Double_Intel (x : Long_Float) return Byte_Buffer;

end Excel_Out;
