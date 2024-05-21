------------------------------------------------------------------------------
--  File:            csv2ada.adb
--  Description:     Converts a CSV (text with Comma Separated Values) input
--                   into an Ada file. The header row will be converted into a
--                   You can specify the separator.
--
--  Syntax:          csv2ada {option} <data.csv
--       or          csv2ada {option}  data.csv
--
--                   Options:
--                     -c : comma is the separator
--                     -s : semicolon is the separator
--                     -t : tab is the separator
--                     -h : has a header row, converted to an enumerated type
--                     -n : name columns
--
--                   For example, the file:
--                       A,B,C
--                       1,2,3
--                       4,5,6
--                       7,8,9
--
--                   is converted, without options, into:
--                       (1 => (A, B, C),
--                        2 => (1, 2, 3),
--                        3 => (4, 5, 6),
--                        4 => (7, 8, 9));
--
--                   with the "-n" option, into:
--                       (1 => (1 => A, 2 => B, 3 => C),
--                        2 => (1 => 1, 2 => 2, 3 => 3),
--                        3 => (1 => 4, 2 => 5, 3 => 6),
--                        4 => (1 => 7, 2 => 8, 3 => 9));
--
--                   with the "-h" option, into:
--                       type Enum is (A, B, C);
--                       (1 => (1, 2, 3),
--                        2 => (4, 5, 6),
--                        3 => (7, 8, 9));
--
--                   with the "-h -n" options, into:
--                       type Enum is (A, B, C);
--                       (1 => (A => 1, B => 2, C => 3),
--                        2 => (A => 4, B => 5, C => 6),
--                        3 => (A => 7, B => 8, C => 9));
--
--  Created:         21-May-2024
--  Author:          Gautier de Montmollin
------------------------------------------------------------------------------

with CSV;

with Ada.Command_Line,
     Ada.Containers.Indefinite_Vectors,
     Ada.Directories,
     Ada.Text_IO,
     Ada.Strings.Fixed;

procedure CSV2Ada is
  use Ada.Command_Line, Ada.Directories, Ada.Text_IO, Ada.Strings, Ada.Strings.Fixed;
  input, output : File_Type;
  lines : Natural := 0;
  data_lines : Natural := 0;
  has_header   : Boolean   := False;
  name_columns : Boolean   := False;
  separator    : Character := ',';
  --  ';', ',' or ASCII.HT

  package Header_Vectors is new Ada.Containers.Indefinite_Vectors (Positive, String);

  header : Header_Vectors.Vector;

begin
  if Argument_Count = 0 then
    Create (output, Out_File, "from_csv.ada");
  else
    declare
      csv_file_name : constant String := Argument (Argument_Count);
      ext : constant String := Extension (csv_file_name);
    begin
      Open (input, In_File, csv_file_name);
      Set_Input (input);
      Create
        (output, Out_File, csv_file_name (csv_file_name'First .. csv_file_name'Last - ext'Length) & "ada");
    end;
  end if;
  --
  --  Process options
  --
  for i in 1 .. Argument_Count loop
    if Argument (i)'Length = 2 and then Argument (i)(1) = '-' then
      case Argument (i)(2) is
        when 'c'    => separator    := ',';
        when 's'    => separator    := ';';
        when 't'    => separator    := ASCII.HT;
        when 'h'    => has_header   := True;
        when 'n'    => name_columns := True;
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
      csv_line : constant String := Get_Line;
      bds  : constant CSV.Fields_Bounds := CSV.Get_Bounds (csv_line, separator);
      procedure Translate_Row is
      begin
        data_lines := data_lines + 1;
        for i in bds'Range loop
          Put
            (output,
             (if i = bds'First then Trim (data_lines'Image, Both) & " => (" else "") &
             (if name_columns then
                (if has_header then header (i) else Trim (i'Image, Both)) &
                " => " else "") &
             CSV.Extract (csv_line, bds, i) &
             (if i = bds'Last then ")" else ", "));
        end loop;
        Put_Line (output, (if End_Of_File then ");" else ","));
      end Translate_Row;
    begin
      lines := lines + 1;
      case lines is
        when 1 =>
          if has_header then
            Put (output, "type Enum is (");
            for i in bds'Range loop
              header.Append (CSV.Extract (csv_line, bds, i));
              Put
                (output, CSV.Extract (csv_line, bds, i) &
                 (if i = bds'Last then ");" else ", "));
            end loop;
            New_Line (output);
            Put_Line (output, "type Data_Row is array (Enum) of Something;");
            Put_Line (output, "type Data_Array is array (Positive range <>) of Data_Row;");
            New_Line (output);
          else
            Put_Line (output, "type Data_Array is array (Positive range <>, Positive range <>) of Something;");
            New_Line (output);
            Put (output, '(');
            Translate_Row;
          end if;
        when 2 =>
          if has_header then
            Put (output, '(');
          else
            Put (output, ' ');
          end if;
          Translate_Row;
        when others =>
          Put (output, ' ');
          Translate_Row;
      end case;
    end;
  end loop;
  Close (output);
  if Is_Open (input) then
    Close (input);
  end if;
end CSV2Ada;
