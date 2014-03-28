-- This test procedure for Excel_Out is in the Ada 95 syntax,
-- for compatibility with a larger number of development systems.
-- With Ada 2005 and later, you can also write "xl.Write(...)" etc. everywhere.
--

with Excel_Out;                         use Excel_Out;

with Ada.Calendar;                      use Ada.Calendar;
with Ada.Numerics.Float_Random;         use Ada.Numerics.Float_Random;
with Ada.Streams.Stream_IO, Ada.Text_IO;

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

  procedure Big_demo(excel_format_choice: Excel_type) is
    xl: Excel_Out_File;
    font_1, font_2, font_3, font_4, font_5, font_6: Font_type;
    fmt_1, fmt_decimal_2, fmt_decimal_0, fmt_4, fmt_5, fmt_6, fmt_cust_num, fmt_8,
    fmt_date_1, fmt_date_2, fmt_date_3: Format_type;
    custom_num, custom_date_num: Number_format_type;
    some_time: constant Time:= Time_Of(2014, 03, 16, (11.0*60.0 + 55.0)* 60.0 + 17.0);
    damier: Natural;
  begin
    Create(xl, "Big [" & Excel_type'Image(excel_format_choice) & "].xls", excel_format_choice);
    -- Some page layout for printing...
    Header(xl, "Big demo");
    Footer(xl, "&D");
    Margins(xl, 1.2, 1.1, 0.9, 0.8);
    Print_Row_Column_Headers(xl);
    Print_Gridlines(xl);
    Page_Setup(xl, fit_height_with_n_pages => 0, orientation => landscape, scale_or_fit => fit);
    --
    Write_default_column_width(xl, 7);
    Write_column_width(xl, 1, 17); -- set to width of n times '0'
    Write_column_width(xl, 2, 11);
    Write_column_width(xl, 5, 11);
    Write_column_width(xl, 14, 0); -- hide this column
    --
    Write_default_row_height(xl, 20);
    -- Write_row_height(xl, 1, 23);   -- header row 1
    -- Write_row_height(xl, 2, 23);   -- header row 2
    Write_row_height(xl, 13, 0);   -- hide this row
    --
    Define_font(xl, "Arial", 9, font_1, regular, blue);
    Define_font(xl, "Courier New", 11, font_2, bold & italic, red);
    Define_font(xl, "Times New Roman", 13, font_3, bold, teal);
    Define_font(xl, "Arial Narrow", 15, font_4, bold);
    Define_font(xl, "Calibri", 15, font_5, bold, dark_red);
    Define_font(xl, "Calibri", 9, font_6);
    --
    Define_number_format(xl, custom_num, "0.000000"); -- 6 decimals
    Define_number_format(xl, custom_date_num, "yyyy\-mm\-dd\ hh:mm:ss"); -- ISO date
    --
    Define_format(xl, font_1, percent_0, fmt_1, centred, right);
    Define_format(xl, font_2, decimal_2, fmt_decimal_2);
    Define_format(xl, font_3, decimal_0_thousands_separator, fmt_decimal_0, centred);
    Define_format(xl, font_4, general,   fmt_4, border => top & bottom);
    Define_format(xl, font_1, percent_2_plus, fmt_5, centred, right);
    Define_format(xl, font_5, general,   fmt_6, border => box);
    Define_format(xl, font_1, custom_num,  fmt_cust_num, centred);
    Define_format(xl, font_6, general, fmt_8);
    Define_format(xl, font_6, dd_mm_yyyy,       fmt_date_1, shaded => True, background_color => yellow);
    Define_format(xl, font_6, dd_mm_yyyy_hh_mm, fmt_date_2, background_color => yellow);
    Define_format(xl, font_6, hh_mm_ss,         fmt_date_3, shaded => True); -- custom_date_num
    --
    Use_format(xl, fmt_4);
    Put(xl, "This is a big demo for Excel Writer / Excel_Out");
    Merge(xl, 6);
    Next(xl);
    Put(xl, "Excel format: " & Excel_type'Image(excel_format_choice));
    Merge(xl, 1);
    New_Line(xl);
    Put(xl, "Version: " & version);
    Merge(xl, 3);
    Next(xl, 4);
    Put(xl, "Ref.: " & reference);

    Use_format(xl, fmt_decimal_2);
    for column in 1 .. 9 loop
      Write(xl, 3, column, Long_Float(column) + 0.5);
    end loop;
    Use_format(xl, fmt_8);
    Put(xl, "  <- = column + 0.5");

    Use_format(xl, fmt_decimal_0);
    for row in 4 .. 7 loop
      for column in 1 .. 9 loop
        damier:= 10 + 990 * ((row + column) mod 2);
        Write(xl, row, column, row * damier + column);
      end loop;
    end loop;
    Use_format(xl, fmt_8);
    Put(xl, "  <- = row * (1000 or 10) + column");

    Use_format(xl, fmt_4);
    for column in 1 .. 20 loop
      Write(xl, 9, column, Character'Val(64 + column) & "");
    end loop;

    Use_format(xl, fmt_6);
    Write(xl, 11, 1, "Calibri font");
    Use_format(xl, fmt_8);
    Write(xl, 11, 4, "First number:");
    Write(xl, 11, 6, Long_Float'First);
    Write(xl, 11, 8, "Last number:");
    Write(xl, 11, 10, Long_Float'Last);
    Write(xl, 11, 12, "Smallest number:");
    Write(xl, 11, 15, (1.0+Long_Float'Model_Epsilon) * Long_Float'Model_Small);
    New_Line(xl);
    -- Date: 2014-03-16 11:55:15
    Use_format(xl, fmt_date_2);
    Put(xl, some_time);
    Use_format(xl, fmt_date_1);
    Put(xl, some_time);
    Use_format(xl, fmt_date_3);
    Put(xl, some_time);
    New_Line(xl);

    for row in 15 .. 300 loop
      Use_format(xl, fmt_1);
      Write(xl, row, 3, Long_Float(row) * 0.01);
      Use_format(xl, fmt_5);
      Put(xl, Long_Float(row-100) * 0.001);
      Use_format(xl, fmt_cust_num);
      Put(xl, Long_Float(row - 15) + 0.123456);
    end loop;
    Close(xl);
  end Big_demo;

  procedure Fancy is
    xl: Excel_Out_File;
    font_title, font_normal, font_normal_grey: Font_type;
    fmt_title, fmt_subtitle, fmt_date, fmt_percent, fmt_amount: Format_type;
    first_day: constant Time:= Time_Of(2014, 03, 28, 9.0*3600.0);
    price, last_price: Long_Float;
    gen: Generator;
  begin
    Create(xl, "Fancy.xls");
    -- Some page layout for printing...
    Header(xl, "Fancy sheet");
    Footer(xl, "&D");
    Margins(xl, 1.2, 1.1, 0.9, 0.8);
    Print_Gridlines(xl);
    Page_Setup(xl, fit_height_with_n_pages => 0, orientation => portrait, scale_or_fit => fit);
    --
    Write_column_width(xl, 1, 15); -- set to width of n times '0'
    Write_column_width(xl, 3, 10); -- set to width of n times '0'
    Define_font(xl, "Calibri", 15, font_title, bold, white);
    Define_font(xl, "Calibri", 10, font_normal);
    Define_font(xl, "Calibri", 10, font_normal_grey, color => grey);
    Define_format(xl, font_title, general, fmt_title, border => bottom, background_color => dark_blue);
    Define_format(xl, font_normal, general, fmt_subtitle, border => bottom);
    Define_format(xl, font_normal, dd_mm_yyyy, fmt_date, background_color => silver);
    Define_format(xl, font_normal, decimal_0_thousands_separator, fmt_amount);
    Define_format(xl, font_normal_grey, percent_2_plus, fmt_percent);
    Use_format(xl, fmt_title);
Write_row_height(xl, 1, 25);
    Put(xl, "Daily Excel Writer stock prices");
close(xl); return;
    Merge(xl, 3);
    New_Line(xl);
    Use_format(xl, fmt_subtitle);
    Put(xl,"Date");
    Put(xl,"Price");
    Put_Line(xl,"Variation %");
    Reset(gen);
    price:= 950.0 + Long_Float(Random(gen)) * 200.0;
    for i in 1..3650 loop
      Use_format(xl, fmt_date);
      Put(xl, first_day + i * Day_Duration'Last);
      Use_format(xl, fmt_amount);
      last_price:= price;
      price:= price * (1.0 + 0.1 * (Long_Float(Random(gen)) - 0.4));
      Put(xl, price);
      Use_format(xl, fmt_percent);
      Put_Line(xl, price / last_price - 1.0);
    end loop;
    Close(xl);
  end Fancy;

  function My_nice_sheet(size: Positive) return String is
    xl: Excel_Out_String;
  begin
    Create(xl);
    Put_Line(xl, "This Excel file is fully created in memory.");
    Put_Line(xl, "It can be stuffed directly into a zip stream,");
    Put_Line(xl, "or sent from a server!");
    Put_Line(xl, "- see ZipTest @ unzip-ada or zip-ada");
    for row in 1 .. size loop
      for column in 1 .. size loop
        Write(xl, row + 5, column, 0.01 + Long_Float(row * column));
      end loop;
    end loop;
    Close(xl);
    return Contents(xl);
  end My_nice_sheet;

  procedure String_demo is
    use Ada.Streams.Stream_IO;
    f: File_Type;
  begin
    Create(f, Out_File, "From_string.xls");
    String'Write(Stream(f), My_nice_sheet(200));
    Close(f);
  end String_demo;

  procedure Speed_test is
    xl: Excel_Out_File;
    t0, t1: Time;
    iter: constant:= 1000;
    size: constant:= 150;
    secs: Long_Float;
  begin
    Create(xl, "Speed_test.xls");
    t0:= Clock;
    for i in 1..iter loop
      declare
        dummy: constant String:= My_nice_sheet(size);
      begin
        if dummy = "" then
          null;
        end if;
      end;
    end loop;
    t1:= Clock;
    secs:= Long_Float(t1-t0);
    Put_Line(xl,
      "Time (seconds) for creating" &
      Integer'Image(iter) & " sheets with" &
      Integer'Image(size) & " x" &
      Integer'Image(size) & " =" &
      Integer'Image(size**2) & " cells"
    );
    Put_Line(xl, secs);
    Put_Line(xl, "Sheets per second");
    Put_Line(xl, Long_Float(iter) / secs);
    Close(xl);
  end Speed_test;

  use Ada.Text_IO;

begin
  Put_Line("Small demo ( -> Small.xls )");
  Small_demo;
  Put_Line("Big demo ( -> Big [...].xls )");
  for f in Excel_type loop
    Big_demo(f);
  end loop;
  Put_Line("Fancy sheet ( -> Fancy.xls )");
  Fancy;
  Put_Line("Excel sheet in a string demo ( -> From_string.xls )");
  String_demo;
  Put_Line("Speed test ( -> Speed_test.xls )");
  Speed_test;
end Excel_Out_Test;
