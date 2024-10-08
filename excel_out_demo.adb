with Excel_Out;

with Ada.Calendar,
     Ada.Characters.Handling,
     Ada.Numerics.Float_Random,
     Ada.Streams.Stream_IO,
     Ada.Text_IO;

procedure Excel_Out_Demo is

  use Excel_Out, Ada.Calendar;

  procedure Small_Demo is
    xl : Excel_Out.Excel_Out_File;
  begin
    xl.Create ("small.xls");
    xl.Put_Line ("This is a small demo for Excel_Out");
    for row in 3 .. 8 loop
      for column in 1 .. 8 loop
        xl.Write (row, column, row * 1000 + column);
      end loop;
    end loop;
    xl.Close;
  end Small_Demo;

  procedure Big_Demo (ef : Excel_Type) is
    xl : Excel_Out_File;
    font_1, font_2, font_3, font_for_title, font_5, font_6 : Font_Type;
    fmt_1, fmt_decimal_2, fmt_decimal_0, fmt_for_title, fmt_5, fmt_boxed, fmt_cust_num, fmt_8,
    fmt_date_1, fmt_date_2, fmt_date_3, fmt_vertical : Format_Type;
    custom_num, custom_date_num : Number_Format_Type;
    --  We test the output of some date (here: 2014-03-16 11:55:17)
    some_time : constant Time := Time_Of (2014, 03, 16, Duration ((11.0 * 60.0 + 55.0) * 60.0 + 17.0));
    damier : Natural;
    first_row : constant := 15;
    use Ada.Characters.Handling;
  begin
    xl.Create ("big [" & To_Lower (ef'Image) & "].xls", ef, Windows_CP_1253);
    xl.Zoom_Level (85, 100);  --  Zoom level 85% (Excel: Ctrl + one bump down with the mouse wheel)
    --  Some page layout for printing...
    xl.Header ("Big demo");
    xl.Footer ("&D");
    xl.Margins (1.2, 1.1, 0.9, 0.8);
    xl.Print_Row_Column_Headers;
    xl.Print_Gridlines;
    xl.Page_Setup (fit_height_with_n_pages => 0, orientation => landscape, scale_or_fit => fit);
    --
    xl.Write_Default_Column_Width (7);
    xl.Write_Column_Width (1, 17);  --  set to width of n times '0'
    xl.Write_Column_Width (2, 11);
    xl.Write_Column_Width (5, 11);
    xl.Write_Column_Width (14, 0);  --  hide this column
    --
    xl.Write_Default_Row_Height (14);
    xl.Write_Row_Height (1, 23);   --  header row 1
    xl.Write_Row_Height (2, 23);   --  header row 2
    xl.Write_Row_Height (9, 23);
    xl.Write_Row_Height (11, 43);
    xl.Write_Row_Height (13, 0);   --  hide this row
    --
    xl.Define_Font ("Arial", 9, font_1, regular, blue);
    xl.Define_Font ("Courier New", 11, font_2, bold & italic, red);
    xl.Define_Font ("Times New Roman", 13, font_3, bold, teal);
    xl.Define_Font ("Arial Narrow", 15, font_for_title, bold);
    xl.Define_Font ("Calibri", 15, font_5, bold, dark_red);
    xl.Define_Font ("Calibri", 9, font_6);
    --
    xl.Define_Number_Format (custom_num, "0.000000");  --  6 decimals
    xl.Define_Number_Format (custom_date_num, "yyyy\-mm\-dd\ hh:mm:ss");  --  ISO date
    --
    xl.Define_Format
      (font           => font_for_title,
       number_format  => general,
       cell_format    => fmt_for_title,
       border         => top & bottom,
       vertical_align => centred);
    --
    xl.Define_Format (font_1, percent_0, fmt_1, centred, right);
    xl.Define_Format (font_2, decimal_2, fmt_decimal_2);
    xl.Define_Format (font_3, decimal_0_thousands_separator, fmt_decimal_0, centred);
    xl.Define_Format (font_1, percent_2_plus, fmt_5, centred, right);
    xl.Define_Format (font_5, general,   fmt_boxed, border => box, vertical_align => centred);
    xl.Define_Format (font_1, custom_num,  fmt_cust_num, centred);
    xl.Define_Format (font_6, general, fmt_8);
    xl.Define_Format (font_6, dd_mm_yyyy,       fmt_date_1, shaded => True, background_color => yellow);
    xl.Define_Format (font_6, dd_mm_yyyy_hh_mm, fmt_date_2, background_color => yellow);
    xl.Define_Format (font_6, hh_mm_ss,         fmt_date_3, shaded => True);  --  custom_date_num
    xl.Define_Format (font_6, general, fmt_vertical, wrap_text => True, text_orient => rotated_90);
    --
    xl.Use_Format (fmt_for_title);
    xl.Put ("This is a big demo for Excel Writer / Excel_Out");
    xl.Merge (6);
    xl.Next;
    xl.Put ("Excel format: " & ef'Image);
    xl.Merge (1);
    xl.New_Line;
    xl.Freeze_Top_Row;
    xl.Put ("Version: " & version);
    xl.Merge (3);
    xl.Next (4);
    xl.Put ("Ref.: " & reference);

    xl.Use_Format (fmt_decimal_2);
    for column in 1 .. 9 loop
      xl.Write (3, column, Long_Float (column) + 0.5);
    end loop;
    xl.Use_Format (fmt_8);
    xl.Put ("  <- = column + 0.5");

    xl.Use_Format (fmt_decimal_0);
    for row in 4 .. 7 loop
      for column in 1 .. 9 loop
        damier := 10 + 990 * ((row + column) mod 2);
        xl.Write (row, column, row * damier + column);
      end loop;
    end loop;
    xl.Use_Format (fmt_8);
    xl.Put ("  <- = row * (1000 or 10) + column");

    xl.Use_Format (fmt_for_title);
    for column in 1 .. 20 loop
      xl.Write (9, column, Character'Val (64 + column) & "");
    end loop;

    xl.Use_Format (fmt_boxed);
    xl.Write (11, 1, "Calibri font");
    xl.Use_Format (fmt_vertical);
    xl.Put ("Wrapped text, rotated 90�");
    xl.Use_Format (fmt_8);
    xl.Write (11, 4, "First number:");
    xl.Write (11, 6, Long_Float'First);
    xl.Write (11, 8, "Last number:");
    xl.Write (11, 10, Long_Float'Last);
    xl.Write (11, 12, "Smallest number:");
    xl.Write (11, 15, (1.0 + Long_Float'Model_Epsilon) * Long_Float'Model_Small);
    xl.Next;
    --  Testing a specific code page (Windows_CP_1253), which was set upon the Create call above.
    xl.Put_Line ("A few Greek letters (alpha, beta, gamma): " &
      Character'Val (16#E1#) & ", " & Character'Val (16#E2#) & ", " & Character'Val (16#E3#)
    );
    --  Date: 2014-03-16 11:55:15
    xl.Use_Format (fmt_date_2);
    xl.Put (some_time);
    xl.Use_Format (fmt_date_1);
    xl.Put (some_time);
    xl.Use_Format (fmt_date_3);
    xl.Put (some_time);
    xl.Use_Default_Format;
    xl.Put (0.0);
    xl.Write_Cell_Comment_at_Cursor ("This is a comment." & ASCII.LF & "Nice, isn't it ?");
    xl.Put (" <- default fmt (general)");
    xl.New_Line;

    for row in first_row .. 300 loop
      xl.Use_Format (fmt_1);
      xl.Write (row, 3, Long_Float (row) * 0.01);
      xl.Use_Format (fmt_5);
      xl.Put (Long_Float (row - 100) * 0.001);
      xl.Use_Format (fmt_cust_num);
      xl.Put (Long_Float (row - 15) + 0.123456);
      xl.Use_Default_Format;
      --  Formulas:
      xl.Put ("=""Formula! ""&E" & Img (row));
      xl.Next (2);
      xl.Put ("= 10000 + E" & Img (row));
      xl.Put ("= 10 * E" & Img (row));
      xl.Next;
      if row = first_row then
        xl.Put (1);
      else
        xl.Put ("=L" & Img (row - 1) & "+L" & Img (row - 2));
      end if;
    end loop;
    xl.Close;
  end Big_Demo;

  procedure Fancy is
    xl : Excel_Out_File;
    font_title, font_normal, font_normal_grey : Font_Type;
    fmt_title, fmt_subtitle, fmt_date, fmt_percent, fmt_amount : Format_Type;
    quotation_day : Time := Time_Of (2023, 03, 28, 9.0 * 3600.0);
    price, last_price : Long_Float;
    use Ada.Numerics.Float_Random;
    gen : Generator;
  begin
    xl.Create ("fancy.xls");
    --  Some page layout for printing...
    xl.Header ("Fancy sheet");
    xl.Footer ("&D");  --  Current date
    xl.Margins (1.2, 1.1, 0.9, 0.8);
    xl.Print_Gridlines;
    xl.Page_Setup (fit_height_with_n_pages => 0, orientation => portrait, scale_or_fit => fit);
    --
    xl.Write_Column_Width (1, 15);  --  set to width of n times '0'
    xl.Write_Column_Width (3, 10);  --  set to width of n times '0'
    xl.Define_Font ("Calibri", 15, font_title, bold, white);
    xl.Define_Font ("Calibri", 10, font_normal);
    xl.Define_Font ("Calibri", 10, font_normal_grey, color => grey);
    xl.Define_Format (font_title, general, fmt_title,
      border => bottom, background_color => dark_blue,
      vertical_align => centred
    );
    xl.Define_Format (font_normal, general, fmt_subtitle, border => bottom);
    xl.Define_Format (font_normal, dd_mm_yyyy, fmt_date, background_color => silver);
    xl.Define_Format (font_normal, decimal_0_thousands_separator, fmt_amount);
    xl.Define_Format (font_normal_grey, percent_2_plus, fmt_percent);
    xl.Use_Format (fmt_title);
    xl.Write_Row_Height (1, 25);
    xl.Put ("Daily Excel Writer stock prices");
    xl.Merge (3);
    xl.New_Line;
    xl.Use_Format (fmt_subtitle);
    xl.Put ("Date");
    xl.Put ("Price");
    xl.Put_Line ("Variation %");
    xl.Freeze_Panes_at_Cursor;
    Reset (gen);
    price := 950.0 + Long_Float (Random (gen)) * 200.0;
    for i in 1 .. 3650 loop
      xl.Use_Format (fmt_date);
      xl.Put (quotation_day);
      quotation_day := quotation_day + Day_Duration'Last;
      xl.Use_Format (fmt_amount);
      last_price := price;
      --  Subtract 0.5 after Random for zero growth / inflation / ...
      price := price * (1.0 + 0.1 * (Long_Float (Random (gen)) - 0.489));
      xl.Put (price);
      xl.Use_Format (fmt_percent);
      xl.Put_Line (price / last_price - 1.0);
    end loop;
    Close (xl);
  end Fancy;

  function My_Nice_Memory_Sheet (size : Positive) return String is
    xl : Excel_Out_String;
  begin
    xl.Create;
    xl.Put_Line ("This Excel file is fully created in memory.");
    xl.Put_Line ("It can be stuffed directly into a zip stream,");
    xl.Put_Line ("or sent from a server!");
    xl.Put_Line ("- see ZipTest @ project Zip-Ada (search ""unzip-ada"" or ""zip-ada""");
    for row in 1 .. size loop
      for column in 1 .. size loop
        xl.Write (row + 5, column, 0.01 + Long_Float (row * column));
      end loop;
    end loop;
    xl.Close;
    return xl.Contents;
  end My_Nice_Memory_Sheet;

  procedure String_Demo is
    use Ada.Streams.Stream_IO;
    f : File_Type;
  begin
    Create (f, Out_File, "from_string.xls");
    String'Write (Stream (f), My_Nice_Memory_Sheet (200));
    Close (f);
  end String_Demo;

  procedure Speed_Test is
    xl : Excel_Out_File;
    t0, t1 : Time;
    iter : constant := 1000;
    size : constant := 150;
    secs : Long_Float;
    dummy_int : Integer := 0;
  begin
    xl.Create ("speed_test.xls");
    t0 := Clock;
    for i in 1 .. iter loop
      declare
        dummy : constant String := My_Nice_Memory_Sheet (size);
      begin
        dummy_int := 0 * dummy_int + dummy'Length;
      end;
    end loop;
    t1 := Clock;
    secs := Long_Float (t1 - t0);
    xl.Put_Line (
      "Time (seconds) for creating" &
      iter'Image & " sheets with" &
      size'Image & " x" &
      size'Image & " =" &
      Integer'Image (size**2) & " cells"
    );
    xl.Put_Line (secs);
    xl.Put_Line ("Sheets per second");
    xl.Put_Line (Long_Float (iter) / secs);
    xl.Close;
  end Speed_Test;

  use Ada.Text_IO;

begin
  Put_Line ("Small demo ----------------------> small.xls");
  Small_Demo;
  Put_Line ("Big demo ------------------------> big [...].xls");
  for ef in BIFF3 .. BIFF4 loop
    Big_Demo (ef);
  end loop;
  Put_Line ("Fancy sheet ---------------------> fancy.xls");
  Fancy;
  Put_Line ("Excel sheet in a string demo ----> from_string.xls");
  String_Demo;
  Put_Line ("Speed test ----------------------> speed_test.xls");
  Speed_Test;
end Excel_Out_Demo;
