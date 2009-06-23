-- This test procedure is in Ada 95 syntax;
-- with Ada 2005 you can write "xl.Write(...)" etc. everywhere.
--

with Excel_Out;                         use Excel_Out;

with Ada.Sequential_IO;

procedure Excel_Out_Test is

  procedure Small_demo is
    xl: Excel_Out_File;
  begin
    Create(xl, "Small.xls");
    Put_Line(xl, "This is a small demo for Excel_Out");
    for row in 3 .. 8 loop
      for column in 1 .. 8 loop
        Write(xl, row, column, row * 1000 + column);
      end loop;
    end loop;
    Close(xl);
  end Small_demo;

  procedure Big_demo is
    xl: Excel_Out_File;
    font_1, font_2, font_3, font_4: Font_Type;
    fmt_1, fmt_2, fmt_3, fmt_4: Format_type;
  begin
    Create(xl, "Big.xls");
    --
    Write_default_column_width(xl, 7);
    Write_column_width(xl, 1, 15); -- set to width of 15 times '0'
    Write_column_width(xl, 14, 0); -- hide this column
    --
    Write_default_row_height(xl, 20);
    Write_row_height(xl, 9, 30);
    Write_row_height(xl, 13, 0);   -- hide this row
    Write_row_height(xl, 100, 30);
    --
    Define_font(xl, "Arial", 10, font_1, regular, blue);
    Define_font(xl, "Courier New", 12, font_2, bold & italic, red);
    Define_font(xl, "Times New Roman", 14, font_3, bold);
    Define_font(xl, "Arial Narrow", 16, font_4, bold);
    --
    Define_format(xl, font_1, percent_0, fmt_1, centred, right);
    Define_format(xl, font_2, decimal_2, fmt_2);
    Define_format(xl, font_3, decimal_0, fmt_3, centred);
    Define_format(xl, font_4, general,   fmt_4, border => top & bottom);
    --
    Use_format(xl, fmt_4);
    Put(xl, "This is a big demo for Excel_Out");
    Jump(xl, 0, 7);
    Put(xl, "Version: " & version);
    Jump_to(xl, 1, 13);
    Put(xl, "Ref.: " & reference);
    Use_format(xl, fmt_2);
    for column in 1 .. 9 loop
      Write(xl, 2, column, Long_Float(column) + 0.5);
    end loop;
    Use_format(xl, fmt_3);
    for row in 4 .. 7 loop
      for column in 1 .. 9 loop
        Write(xl, row, column, row * 1000 + column);
      end loop;
    end loop;
    Use_format(xl, fmt_4);

    for column in 1 .. 20 loop
      Write(xl, 9, column, Character'Val(64 + column) & "");
    end loop;
    Use_format(xl, fmt_1);
    for row in 13 .. 300 loop
      Write(xl, row, 3, Long_Float(row) * 0.01);
    end loop;
    Close(xl);
  end Big_demo;

  function My_nice_sheet return String is
    xl: Excel_Out_String;
  begin
    Create(xl);
    Put_Line(xl, "This Excel file is fully created in memory.");
    Put_Line(xl, "It can be stuffed directly into a zip stream,");
    Put_Line(xl, "or sent from a server!");
    Put_Line(xl, "- see ZipTest @ unzip-ada or zip-ada");
    Close(xl);
    return Contents(xl);
  end My_nice_sheet;

  procedure String_demo is
    package CIO is new Ada.Sequential_IO(Character);
    xls: constant String:= My_nice_sheet;
    use CIO;
    f: File_Type;
  begin
    Create(f, Out_File, "From_string.xls");
    for i in xls'Range loop
      Write(f, xls(i));
    end loop;
    Close(f);
  end String_demo;

begin
  Small_demo;
  Big_demo;
  String_demo;
end Excel_Out_Test;
