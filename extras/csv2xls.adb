------------------------------------------------------------------------------
--  File:            CSV2XLS.adb
--  Description:     Converts a CSV (text with Comma Separated Values) input
--                   into an Excel file. You can specify the separator.
--                   E.g. If you open in Excel a CSV with semicolons as
--                   separators on a PC with comma being the separator,
--                   and then apply "Text to columns", the eventual commas in
--                   the text will already have been used as separators and
--                   you will end up with a total mess. CSV2XLS prevents this
--                   issue.
--  Syntax:          csv2xls {option} <data.csv
--       or          csv2xls {option} data.csv
--                   Options:
--                     -c : comma is the separator
--                     -s : semicolon is the separator
--                     -t : tab is the separator
--                     -f : freeze top row (header line)
--  Created:         29-Apr-2014
--  Author:          Gautier de Montmollin
------------------------------------------------------------------------------

with CSV;
with Excel_Out;
with Ada.Command_Line, Ada.Directories, Ada.Text_IO, Ada.Strings.Fixed;

procedure CSV2XLS is
  use Ada.Command_Line, Ada.Directories, Ada.Text_IO, Ada.Strings, Excel_Out;
  input : File_Type;
  xl : Excel_Out_File;
  first : Boolean := True;
  separator : Character := ',';
  --  ';', ',' or ASCII.HT
begin
  if Argument_Count = 0 then
    Create (xl, "From_CSV.xls");
  else
    declare
      csv_file_name : constant String := Argument (Argument_Count);
      ext : constant String := Extension (csv_file_name);
    begin
      Open (input, In_File, csv_file_name);
      Set_Input (input);
      Create (xl, csv_file_name (csv_file_name'First .. csv_file_name'Last - ext'Length) & "xls");
    end;
  end if;
  --
  --  Process options
  --
  for i in 1 .. Argument_Count loop
    if Argument (i)'Length = 2 and then Argument (i)(1) = '-' then
      case Argument (i)(2) is
        when 'c' =>
          separator := ',';
        when 's' =>
          separator := ';';
        when 't' =>
          separator := ASCII.HT;
        when 'f' =>
          Freeze_Top_Row (xl);
        when others =>
          null;
      end case;
    end if;
  end loop;
  --
  --  Process the CSV file
  --
  while not End_Of_File loop
    declare
      line : constant String := Get_Line;
      bds : constant CSV.Fields_Bounds := CSV.Get_Bounds (line, separator);
    begin
      if first then
        first := False;
      end if;
      for i in bds'Range loop
        Put (xl, Ada.Strings.Fixed.Trim (CSV.Extract (line, bds, i), Both));
      end loop;
    end;
    New_Line (xl);
  end loop;
  Close (xl);
  if Is_Open (input) then
    Close (input);
  end if;
end CSV2XLS;
