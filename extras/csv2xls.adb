------------------------------------------------------------------------------
--  File:            CSV2XLS.adb
--  Description:     Converts a CSV (text with Comma Separated Values) input
--                   into Excel file. You can specify the separator.
--                   E.g. If you open in Excel a CSV with semicolons as
--                   separators on a PC with comma being the separator,
--                   and then apply "Text to columns", the eventual commas in
--                   the text will already have been used as separators and
--                   you will end up with a total mess. CSV2XLS prevents this
--                   issue.
--  Syntax:          csv2xls <data.csv
--       or          csv2xls data.csv
--  Date / Version:  29-Apr-2014
--  Author:          Gautier de Montmollin
------------------------------------------------------------------------------

with CSV;
with Excel_Out;
with Ada.Command_Line, Ada.Directories, Ada.Text_IO, Ada.Strings.Fixed;

procedure CSV2XLS is
  use Ada.Command_Line, Ada.Directories, Ada.Text_IO, Ada.Strings, Excel_Out;
  input: File_Type;
  xl: Excel_Out_File;
  first: Boolean:= True;
  separator: constant Character := ',';
  -- ';', ',' or ASCII.HT
begin
  if Argument_Count = 0 then
    Create(xl, "From_CSV.xls");
  else
    declare
      name: constant String:= Argument(1);
      ext: constant String:= Extension(name);
    begin
      Open(input, In_File, name);
      Set_Input(input);
      Create(xl, name(name'First..name'Last-ext'Length) & "xls");
    end;
  end if;
  while not End_Of_File loop
    declare
      line: constant String:= Get_Line;
      bds: constant CSV.Fields_Bounds:= CSV.Get_Bounds( line, separator );
    begin
      if first then
        first:= False;
      end if;
      for i in bds'Range loop
        Put(xl,
          Ada.Strings.Fixed.Trim(
            CSV.Unquote(CSV.Extract(line, bds, i)), Both
          )
        );
      end loop;
    end;
    New_Line(xl);
  end loop;
  Close(xl);
  if Is_Open(input) then
    Close(input);
  end if;
end CSV2XLS;
