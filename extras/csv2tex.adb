------------------------------------------------------------------------------
--  File:            csv2tex.adb
--  Description:     Converts CSV (text with Comma Separated Values) input
--                   into LaTeX array output. NB: the special characters
--                   like '%', '\', '&', '{',... should be translated before!
--                   CSV is "the" ASCII format for Lotus 1-2-3 and MS Excel
--
--  Created:         22-Apr-2003
--  Author:          Gautier de Montmollin
------------------------------------------------------------------------------

with Ada.Text_IO, Ada.Strings.Fixed;
with CSV; -- replaces CSV_Parser

procedure CSV2TeX is
  use Ada.Text_IO, Ada.Strings;
  first : Boolean := True;
  separator : constant Character := ';';
  --  ';', ',' or ASCII.HT
begin
  while not End_Of_File loop
    declare
      csv_line : constant String := Get_Line;
      bds : constant CSV.Fields_Bounds := CSV.Get_Bounds (csv_line, separator);
    begin
      if first then
        Put_Line ("% Array translated by CSV2TeX");
        Put_Line ("% Check http://excel-writer.sourceforge.net/ ,");
        Put_Line ("% in the ./extras directory");
        Put ("\begin{array}{");
        for i in bds'Range loop
          Put ('c');
          if i < bds'Last then
            Put ('|');
          end if;
        end loop;
        Put_Line ("} % array description");
        first := False;
      end if;
      for i in bds'Range loop
        Put (Ada.Strings.Fixed.Trim (CSV.Extract (csv_line, bds, i), Both));
        if i < bds'Last then
          Put ("&");
        end if;
      end loop;
    end;
    Put_Line ("\\");
  end loop;
  Put_Line ("\end{array}");
end CSV2TeX;
