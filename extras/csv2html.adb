------------------------------------------------------------------------------
--  File:            CSV2HTML.adb
--  Description:     Converts CSV (text with Comma Separated Values) input
--                   into HTML array output.
--                   CSV is "the" ASCII format for Lotus 1-2-3 and MS Excel
--  Created:         11-Nov-2002
--  Syntax (example) csv2html <demo_data.csv >demo_data.html
--  Author:          Gautier de Montmollin
------------------------------------------------------------------------------

with Ada.Text_IO, Ada.Strings.Fixed;
with CSV; -- replaces CSV_Parser

procedure CSV2HTML is

  function Windows8bit_to_HTML (c : Character) return String is
  begin
    case c is
      when 'à' => return "&agrave;";
      when 'â' => return "&acirc;";
      when 'ä' => return "&auml;";
      when 'Ä' => return "&Auml;";
      when 'é' => return "&eacute;";
      when 'É' => return "&Eacute;";
      when 'è' => return "&egrave;";
      when 'È' => return "&Egrave;";
      when 'î' => return "&icirc;";
      when 'ô' => return "&ocirc;";
      when 'ö' => return "&ouml;";
      when 'Ö' => return "&Ouml;";
      when 'ß' => return "&szlig;";
      when 'û' => return "&ucirc;";
      when 'ü' => return "&uuml;";
      when 'Ü' => return "&Uuml;";
      when '’' => return "'";
      when '´' => return "'";
      when ' ' => return "&nbsp;";
      when others =>
        return (1 => c);
    end case;
  end Windows8bit_to_HTML;

  function Windows8bit_to_HTML (s : String) return String is
  begin
    if s = "" then
      return "";
    else
      return
        Windows8bit_to_HTML (s (s'First)) &
        Windows8bit_to_HTML (s (s'First + 1 .. s'Last));
    end if;
  end Windows8bit_to_HTML;

  line_count : Natural := 0;
  special_sport : constant Boolean := False;
  separator : constant Character := ',';
  --  ';', ',' or ASCII.HT

  use Ada.Text_IO, Ada.Strings;

begin
  Put_Line ("<!--  Array translated by CSV2HTML                  !--> ");
  Put_Line ("<!--  Check http://excel-writer.sourceforge.net/ ,  !--> ");
  Put_Line ("<!--  in the ./extras directory                     !--> ");
  Put_Line ("<table border=2><tr>");
  --
  while not End_Of_File loop
    line_count := line_count + 1;
    declare
      csv_line : constant String := Get_Line;
      bds  : constant CSV.Fields_Bounds := CSV.Get_Bounds (csv_line, separator);
    begin
      for i in bds'Range loop
        Put ("<td>");
        --  Trim blanks on both sides (7-Dec-2003)
        Put (
          Windows8bit_to_HTML (
            Ada.Strings.Fixed.Trim (CSV.Extract (csv_line, bds, i), Both)
          )
        );
        Put ("</td>");
      end loop;
    end;
    Put_Line ("</td></tr>");
    --
    if special_sport then
      case line_count is
        when 1 => Put ("<tr bgcolor=#fbee33>");  --  Gold
        when 2 => Put ("<tr bgcolor=#c3c3c3>");  --  Silver
        when 3 => Put ("<tr bgcolor=#c89544>");  --  Bronze
        when others => Put ("<tr>");
      end case;
    else
      Put ("<tr>");
    end if;
    --
  end loop;
  Put_Line ("</tr></table>");
end CSV2HTML;
