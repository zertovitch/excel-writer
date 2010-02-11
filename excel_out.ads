-------------------------------------------------------------------------------------
--
-- EXCEL_OUT - A low level package to write Microsoft Excel (*) files
--
-- Pure Ada 95 code, 100% portable: OS-, CPU- and compiler- independent.
--

-- Legal licensing note:

--  Copyright (c) 2009..2010 Gautier de Montmollin

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

-- NB: this is the MIT License, as found 12-Sep-2007 on the site
-- http://www.opensource.org/licenses/mit-license.php

-- Derived from ExcelOut @ http://www.modula2.org/projects/excelout.php
-- by Frank Schoonjans - thanks!

-- (*) All Trademarks mentioned are properties of their respective owners.
-------------------------------------------------------------------------------------
--
--  Follow these steps to create a stream with the contents of an Excel file:
--
--  1. Create
--
--  2. Optional settings, before any data output:
--     | Write_default_column_width
--     | Write_column_width for specific columns
--     | Write_default_row_height
--     | Write_row_height for specific rows
--     | Define_font, then Define_format
--
--  3. | Write(xl, row, column, data): row by row, column by column
--     | Put(xl, data)               : same, but column is auto-incremented
--     | New_Line(xl),...            : other "Text_IO"-like (full list below)
--     | Use_format, infuences the format of data written next
--
--  4. Close
--
--  5. (Excel_Out_String only) function Contents returns the full .xls
--------------------------------------------------------------------------
--
-- Main changes:
-- ============
--
-- 05: 10-Feb-2010: - added 'width' and 'base' optional parameters
--                      to Put(xl, int), to facilitate transition
--                      from Ada.Text_IO.* to Excel_Out
--                  - added function Is_Open(xl : in Excel_Out_File)
--
-- 03: 15-Feb-2009: - data stream can by any; supplied:
--                      Excel_Out_File, Excel_Out_String
--                  - added "Text_IO"-like Put, Put_Line, New_Line,...
-- 02: 14-Feb-2009: - row/column coordinates are 1-based (they have to!)
--                  - added horizontal alignment and cell borders
-- 01: 13-Feb-2009: 1st release
-- 00: 11-Feb-2009: translation from original ExcelOut in Modula-2

with Ada.Streams.Stream_IO;
with Ada.Strings.Unbounded;             use Ada.Strings.Unbounded;
with Ada.Text_IO;

package Excel_Out is

  ----------------------------------------------------------------
  -- The Excel output stream root type. Cannot be used as such. --
  -- From this package, you can use the following types:        --
  --  * Excel_Out_File    : output in a file                    --
  --  * Excel_Out_String  : output in a string                  --
  -- Of course you can define your own derived types.           --
  ----------------------------------------------------------------

  type Excel_Out_Stream is abstract tagged private;

  type Excel_type is (
    BIFF2     -- Excel 2.x: the oldest & easiest format
              --            - and necessarily the most compatible
    -- BIFF3,    -- Excel 3.0
    -- BIFF4,    -- Excel 4.0
    -- BIFF5,    -- Excel 5.0 to 7.0
    -- BIFF8     -- Excel 8.0 to 11.0
  );

  Default_Excel_type: constant Excel_type:= BIFF2;

  ----------------------------------
  -- (2) Before any cell content: --
  ----------------------------------

  -- * The column width unit is as it appears in Excel when you resize a column.
  --     It is the width of a '0' in a standard font.
  procedure Write_default_column_width(xl : Excel_Out_Stream; width : Positive);
  procedure Write_column_width(xl : Excel_Out_Stream; column: Positive; width: Natural);

  -- * The row height unit is in font points, as appearing when you
  --     resize a row in Excel.
  procedure Write_default_row_height(xl: Excel_Out_Stream; height: Positive);
  procedure Write_row_height(xl : Excel_Out_Stream; row: Positive; height : Natural);

  -- * Fonts are user-defined, one is predefined: Default_font
  type Font_type is private;

  function Default_font(xl: Excel_Out_Stream) return Font_type;
  -- Arial 10, regular, "automatic" color

  type Color_type is
    (automatic, black, white, red, green, blue, yellow, magenta, cyan);

  type Font_style is private;

  -- for combining styles (e.g.: bold & underlined):
  function "&"(a,b: Font_style) return Font_style;

  regular     : constant Font_style;
  italic      : constant Font_style;
  bold        : constant Font_style;
  bold_italic : constant Font_style;
  underlined  : constant Font_style;
  struck_out  : constant Font_style;
  shadowed    : constant Font_style;
  condensed   : constant Font_style;
  extended    : constant Font_style;

  procedure Define_font(
    xl           : in out Excel_Out_Stream;
    font_name    :        String;
    height       :        Positive;
    font         :    out Font_type;
    -- optional:
    style        :        Font_style:= regular;
    color        :        Color_type:= automatic
  );

  -- * Format: a [cell] format is a combination of a font, a number format
  -- (and other optional things to come...)
  -- Formats are user-defined, one is predefined: Default_format
  type Format_type is private;

  function Default_format(xl: Excel_Out_Stream) return Format_type;
  -- What you get when creating a new sheet in Excel: Default_font,...

  type Number_format_type is
    ( general,
      decimal_0, decimal_2,
      decimal_0_thousands_separator,  -- 1'234'000
      decimal_2_thousands_separator,  -- 1'234'000.00
      percent_0, percent_2,           --  3%, 0%, -4%
      percent_0_plus, percent_2_plus, -- +3%, 0%, -4%
      scientific
    );
  -- Number formats are all predefined: must be listed as built-in, grouped;
  -- A number format working on Excel with certain regional settings may not
  -- work on Excel (even the same) with other regional settings!
  -- Hence the limited choice above.

  type Horizontal_alignment is (
    general, to_left, centred, to_right, filled,
    justified, centred_across_selection, -- (BIFF4-BIFF8)
    distributed -- (BIFF8, Excel 10.0 and later only)
  );

  type Cell_border is private;

  -- for combining borders (e.g.: left & top):
  function "&"(a,b: Cell_border) return Cell_border;

  no_border : constant Cell_border;
  left      : constant Cell_border;
  right     : constant Cell_border;
  top       : constant Cell_border;
  bottom    : constant Cell_border;

  procedure Define_format(
    xl           : in out Excel_Out_Stream;
    font         : in     Font_type;   -- given by Define_font
    number_format: in     Number_format_type;
    format       :    out Format_type;
    -- optional:
    horiz_align  : in     Horizontal_alignment:= general;
    border       : in     Cell_border:= no_border;
    shaded       : in     Boolean:= False
  );

  ------------------------
  -- (3) Cell contents: --
  ------------------------

  -- NB: you need to write row by row (ascending index), column by
  --     column (ascending index), otherwise Excel issues a protest

  procedure Write(xl: in out Excel_Out_Stream; r,c : Positive; num : Long_Float);
  procedure Write(xl: in out Excel_Out_Stream; r,c : Positive; num : Integer);
  procedure Write(xl: in out Excel_Out_Stream; r,c : Positive; str : String);
  procedure Write(xl: in out Excel_Out_Stream; r,c : Positive; str : Unbounded_String);

  -- Ada.Text_IO - like. No need to specify row & column each time
  procedure Put(xl: in out Excel_Out_Stream; num : Long_Float);
  procedure Put(xl    : in out Excel_Out_Stream;
                num   : in Integer;
                width : in Ada.Text_IO.Field := 0; -- ignored
                base  : in Ada.Text_IO.Number_Base := 10
            );
  procedure Put(xl: in out Excel_Out_Stream; str : String);
  procedure Put(xl: in out Excel_Out_Stream; str : Unbounded_String);
  --
  procedure Put_Line(xl: in out Excel_Out_Stream; num : Long_Float);
  procedure Put_Line(xl: in out Excel_Out_Stream; num : Integer);
  procedure Put_Line(xl: in out Excel_Out_Stream; str : String);
  procedure Put_Line(xl: in out Excel_Out_Stream; str : Unbounded_String);
  --
  procedure New_Line(xl: in out Excel_Out_Stream);
  -- Relative / absolute jumps
  procedure Jump(xl: in out Excel_Out_Stream; rows, columns: Natural);
  procedure Jump_to(xl: in out Excel_Out_Stream; row, column: Positive);


  -- Cells written after Use_format will be using the given format,
  -- defined by Define_format.
  procedure Use_format(
    xl           : in out Excel_Out_Stream;
    format       : in     Format_type
  );
  procedure Use_default_format(xl: in out Excel_Out_Stream);

  Excel_stream_not_created,
  Excel_stream_not_closed,
  Decreasing_row_index,
  Decreasing_column_index,
  Row_out_of_range,
  Column_out_of_range : exception;

  -----------------------------------------------------------------
  -- Here, the derived stream types pre-defined in this package. --
  -----------------------------------------------------------------
  -- * Output to file:

  type Excel_Out_File is new Excel_Out_Stream with private;

  procedure Create(
    xl        : in out Excel_Out_File;
    file_name :        String;
    format    :        Excel_type:= Default_Excel_type
  );

  procedure Close(xl : in out Excel_Out_File);

  function Is_Open(xl : in Excel_Out_File) return Boolean;

  -- * Output to string:

  type Excel_Out_String is new Excel_Out_Stream with private;

  procedure Create(
    xl        : in out Excel_Out_String;
    format    :        Excel_type:= Default_Excel_type
  );

  procedure Close(xl : in out Excel_Out_String);

  function Contents(xl: Excel_Out_String) return String;

  --------------------------------------------------------------
  -- Information about this package - e.g. for an "about" box --
  --------------------------------------------------------------

  version   : constant String:= "05";
  reference : constant String:= "xx-Feb-2010";
  web       : constant String:= "http://excel-writer.sf.net/";
  -- hopefully the latest version is at that URL...  ---^

  ----------------------------------
  -- End of the part for the user --
  ----------------------------------

  -- Set the index on the stream
  procedure Set_Index (xl: in out Excel_Out_Stream;
                       to: Ada.Streams.Stream_IO.Positive_Count)
  is abstract;

  -- Return the index of the stream
  function Index (xl: Excel_Out_Stream) return Ada.Streams.Stream_IO.Count
  is abstract;

private

  ----------------------------------------
  -- Raw Streams, with 'Read and 'Write --
  ----------------------------------------

  type XL_Raw_Stream_Class is access all Ada.Streams.Root_Stream_Type'Class;

  type Font_type is new Natural;
  type Format_type is new Natural;

  type XF_Range is range 0..62; -- after 62 we'd need to use an IXFE (5.62)
  max_font  : constant:= 62;
  max_format: constant:= 62;

  type Excel_Out_Stream is abstract tagged record
    xl_stream  : XL_Raw_Stream_Class;
    format     : Excel_type:= Default_Excel_type;
    dimrecpos  : Ada.Streams.Stream_IO.Positive_Count;
    maxcolumn  : Integer:= 0;  -- 0-based
    maxrow     : Integer:= 0;  -- 0-based
    fonts      : Integer:= -1; -- [-1..max_font]
    xfs        : Integer:= -1; -- [-1..XF_Range'Last]
    xf_in_use  : XF_Range:= 0;
    def_font   : Font_type;
    def_fmt    : Format_type;
    is_created : Boolean:= False;
    is_closed  : Boolean:= False;
    curr_row   : Positive:= 1;
    curr_col   : Positive:= 1;
  end record;

  type Font_style_single is
    (bold_single,
     italic_single,
     underlined_single,
     struck_out_single,
     outlined_single,
     shadowed_single,
     condensed_single,
     extended_single);

  type Font_style is array(Font_style_single) of Boolean;

  function "&"(a,b: Font_style) return Font_style
    renames "or"; -- predefined for sets (=array of Boolean)

  regular     : constant Font_style:= (others => False);
  italic      : constant Font_style:= (italic_single => True, others => False);
  bold        : constant Font_style:= (bold_single => True, others => False);
  bold_italic : constant Font_style:= bold or italic;
  underlined  : constant Font_style:= (underlined_single => True, others => False);
  struck_out  : constant Font_style:= (struck_out_single => True, others => False);
  shadowed    : constant Font_style:= (shadowed_single => True, others => False);
  condensed   : constant Font_style:= (condensed_single => True, others => False);
  extended    : constant Font_style:= (extended_single => True, others => False);

  type Cell_border_single is
    (left_single,
     right_single,
     top_single,
     bottom_single);

  type Cell_border is array(Cell_border_single) of Boolean;

  function "&"(a,b: Cell_border) return Cell_border
    renames "or"; -- predefined for sets (=array of Boolean)

  no_border : constant Cell_border:= (others => False);
  left      : constant Cell_border:= (left_single => True, others => False);
  right     : constant Cell_border:= (right_single => True, others => False);
  top       : constant Cell_border:= (top_single => True, others => False);
  bottom    : constant Cell_border:= (bottom_single => True, others => False);

  ----------------------
  -- Output to a file --
  ----------------------

  type XL_file_acc is
    access Ada.Streams.Stream_IO.File_Type;

  type Excel_Out_File is new Excel_Out_Stream with record
    xl_file   : XL_file_acc:= null; -- access to the "physical" Excel file
  end record;

  -- Set the index on the file
  procedure Set_Index (xl: in out Excel_Out_File;
                       To: Ada.Streams.Stream_IO.Positive_Count);

  -- Return the index of the file
  function Index (xl: Excel_Out_File) return Ada.Streams.Stream_IO.Count;

  ------------------------
  -- Output to a string --
  ------------------------
  -- Code reused from Zip_Streams

  --- *** We define here a complete in-memory stream:
  type Unbounded_Stream is new Ada.Streams.Root_Stream_Type with
    record
      Unb : Ada.Strings.Unbounded.Unbounded_String;
      Loc : Integer := 1;
    end record;

  -- Read data from the stream.
  procedure Read
    (Stream : in out Unbounded_Stream;
     Item   : out Ada.Streams.Stream_Element_Array;
     Last   : out Ada.Streams.Stream_Element_Offset);

  -- write data to the stream, starting from the current index.
  -- Data will be overwritten from index is already available.
  procedure Write
    (Stream : in out Unbounded_Stream;
     Item   : Ada.Streams.Stream_Element_Array);

  -- Set the index on the stream
  procedure Set_Index (S : access Unbounded_Stream; To : Positive);

  -- returns the index of the stream
  function Index (S: access Unbounded_Stream) return Integer;

  --- ***

  type Unbounded_Stream_Acc is access Unbounded_Stream;

  type Excel_Out_String is new Excel_Out_Stream with record
    xl_memory: Unbounded_Stream_Acc;
  end record;

  -- Set the index on the Excel string stream
  procedure Set_Index (xl: in out Excel_Out_String;
                       To: Ada.Streams.Stream_IO.Positive_Count);

  -- Return the index of the Excel string stream
  function Index (xl: Excel_Out_String) return Ada.Streams.Stream_IO.Count;

end Excel_Out;
