------------------------------------------------------------------------------
--  File:            paypal.adb
--
--  Description:     Converts a PayPal activity report in CSV format
--                   ( https://www.paypal.com/reports/dlog )
--                   into an Excel file, with a section per currency.
--
--  Syntax:          paypal report.csv {other_report.csv}
--
--  Created:         06-Jan-2024
--
--  Author:          Gautier de Montmollin
------------------------------------------------------------------------------

with CSV;

with Excel_Out;

with Ada.Command_Line,
     Ada.Containers.Indefinite_Ordered_Sets,
     Ada.Directories,
     Ada.Text_IO;

procedure PayPal is

  procedure Process (csv_file_name : String) is

    use Ada.Directories, Ada.Text_IO, Excel_Out;

    input : File_Type;
    xl : Excel_Out_File;
    separator : constant Character := ',';

    type Col_Type is
      (Date, Time, Time_Zone,
       Name, Transaction_Type, Status,
       Currency, Amount, Receipt_ID, Balance);

    package String_Ordered_Sets is
      new Ada.Containers.Indefinite_Ordered_Sets (Element_Type => String);

    currency_set : String_Ordered_Sets.Set;

    fmt_title, fmt_subtitle, fmt_header, fmt_money, fmt_normal : Format_type;

    procedure Put_Header is
    begin
      xl.Use_format (fmt_header);
      xl.Put ("Date");
      xl.Put ("Time");
      xl.Put ("TimeZone");
      xl.Put ("Name");
      xl.Put ("Type");
      xl.Put ("Status");
      xl.Put ("Currency");
      xl.Put ("Amount");
      xl.Put ("Receipt ID");
      xl.Put ("Balance");
      xl.New_Line;
    end Put_Header;

    ext : constant String := Extension (csv_file_name);

    prefix : constant String :=
      csv_file_name (csv_file_name'First .. csv_file_name'Last - ext'Length - 1);

    font_title, font_subtitle, font_normal : Font_type;
  begin
    xl.Create (prefix & ".xls");
    xl.Header ("PayPal activity report");
    xl.Footer ("&D");  --  Current date
    xl.Print_Gridlines;
    xl.Page_Setup
      (fit_height_with_n_pages => 0,
       orientation             => landscape,
       scale_or_fit            => fit);

    for ct in Col_Type loop
      xl.Write_column_width
        (1 + Col_Type'Pos (ct),
         (case ct is
            when Date             => 10,
            when Time             => 10,
            when Time_Zone        => 10,
            when Name             => 25,
            when Transaction_Type => 35,
            when Status           => 10,
            when Currency         => 8,
            when Amount           => 9,
            when Receipt_ID       => 10,
            when Balance          => 9));
    end loop;

    xl.Define_Font ("Calibri", 10, font_normal);
    xl.Define_Font ("Calibri", 12, font_subtitle, bold);
    xl.Define_Font ("Calibri", 15, font_title, bold);
    xl.Define_Format (font_normal,   general, fmt_normal);
    xl.Define_Format (font_subtitle, general, fmt_subtitle);
    xl.Define_Format (font_title,    general, fmt_title);
    xl.Define_Format (font_normal,   general, fmt_header, border => bottom);
    xl.Define_Format (font_normal,   decimal_2_thousands_separator, fmt_money);
    xl.Use_format (fmt_normal);

    --
    --  Map the currencies
    --
    Open (input, In_File, csv_file_name);
    Skip_Line (input);  --  Skip header
    while not End_Of_File (input) loop
      declare
        line     : constant String := Get_Line (input);
        bds      : constant CSV.Fields_Bounds := CSV.Get_Bounds (line, separator);
        line_ccy : constant String := CSV.Extract (line, bds, Col_Type'Pos (Currency) + 1);
      begin
        currency_set.Include (line_ccy);
      end;
    end loop;
    Close (input);

    --
    --  Output the report, grouped by currency
    --
    xl.Use_format (fmt_title);
    xl.Put_Line ("PayPal activity report: " & csv_file_name);
    xl.New_Line;
    for ccy of currency_set loop
      xl.Use_format (fmt_subtitle);
      xl.Put_Line ("Currency: " & ccy);
      xl.New_Line;
      Put_Header;
      Open (input, In_File, csv_file_name);
      Skip_Line (input);  --  Skip header
      while not End_Of_File (input) loop
        declare
          line     : constant String := Get_Line (input);
          bds      : constant CSV.Fields_Bounds := CSV.Get_Bounds (line, separator);
          line_ccy : constant String := CSV.Extract (line, bds, Col_Type'Pos (Currency) + 1);
        begin
          if ccy = line_ccy then
            for col in bds'Range loop
              declare
                cell : String := CSV.Extract (line, bds, col);
              begin
                if Col_Type'Val (col - 1) in Amount | Balance then
                  --  Convert commas into dots for monetary values
                  for c of cell loop
                    if c = ',' then c := '.'; end if;
                  end loop;
                  xl.Use_format (fmt_money);
                  xl.Put (Long_Float'Value (cell));
                else
                  xl.Use_format (fmt_normal);
                  xl.Put (cell);
                end if;
              end;
            end loop;
            xl.New_Line;
          end if;
        end;
      end loop;
      xl.New_Line (2);
      Close (input);
    end loop;

    xl.Close;
  end Process;

  use Ada.Command_Line, Ada.Text_IO;

begin
  if Argument_Count = 0 then
    Put_Line (Current_Error, "Syntax: paypal report.csv {other_report.csv}");
    delay 5.0;
  else
    for a in 1 .. Argument_Count loop
      Process (Argument (a));
    end loop;
  end if;
end PayPal;
