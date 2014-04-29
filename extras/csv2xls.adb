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
--  Date / Version:  29-Apr-2014
--  Author:          Gautier de Montmollin
------------------------------------------------------------------------------

with CSV;
with Excel_Out;
with Ada.Text_IO, Ada.Strings.Fixed;

procedure CSV2XLS is
  use Ada.Text_IO, Ada.Strings, Excel_Out;
  xl: Excel_Out_File;
  first: Boolean:= True;
  separator: constant Character := ',';
  -- ';', ',' or ASCII.HT
begin
  Create(xl, "translated.xls");
  while not End_Of_File(Standard_Input) loop
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
end CSV2XLS;
