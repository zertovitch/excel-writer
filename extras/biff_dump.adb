-- Dump the contents of a file in BIFF format

with Ada.Command_Line;                  use Ada.Command_Line;
with Ada.Text_IO;                       use Ada.Text_IO;
with Ada.Integer_Text_IO;               use Ada.Integer_Text_IO;
with Ada.Sequential_IO;
with Interfaces;                        use Interfaces;

procedure BIFF_Dump is

  package BIO is new Ada.Sequential_IO(Unsigned_8);
  use BIO;

  f: BIO.File_Type;

  code, length: Integer;

  function in16 return Integer is
    b1,b2: Unsigned_8;
  begin
    Read(f,b1);
    Read(f,b2);
    return Integer(b1) + Integer(b2) * 256;
  end in16;

  row: constant:= 16#0008#;

  b: Unsigned_8;

begin
  if Argument_Count = 0 then
    Open(f, In_File, "big.xls");
  else
    Open(f, In_File, Argument(1));
  end if;
  while not End_of_File(f) loop
    code  := in16;
    length:= in16;
    Put(code, base => 16);
    Put(length);
    Put("    ");
    case code is
      when 16#0009# => Put("BOF - Beginning of File");
      when 16#000A# => Put("EOF - End of File");
      when 16#0000# => Put("DIMENSION");
      when 16#000D# => Put("CALCMODE");
      when 16#000F# => Put("REFMODE");
      when 16#0022# => Put("DATEMODE");
      when 16#0042# => Put("CODEPAGE");
      when 16#0024# => Put("COLWIDTH");
      when 16#0055# => Put("DEFCOLWIDTH");
      when 16#0025# => Put("DEFAULTROWHEIGHT");
      when row      => Put("ROW");
      when 16#001E# => Put("FORMAT");
      when 16#001F# => Put("BUILTINFMTCOUNT");
      when 16#0031# => Put("FONT");
      when 16#0045# => Put("FONTCOLOR");
      when 16#0001# => Put("BLANK");
      when 16#0002# => Put("INTEGER");
      when 16#0003# => Put("NUMBER");
      when 16#0004# => Put("LABEL");
      when 16#0043# => Put("XF - Extended Format");
      when 16#0019# => Put("WINDOWPROTECT");
      when 16#0040# => Put("BACKUP");
      when others =>   Put("- ??? -");
    end case;
    case code is
      when row =>
        Put("  row="); Put(in16,0);
        Put(" col1="); Put(in16,0);
        Put(" col2="); Put(in16,0);
        Put(" height="); Put(in16,0);
        for i in 1..5 loop
          Read(f,b);
        end loop;
      when 1..4 =>
        Put("  row="); Put(in16,0);
        Put("  col="); Put(in16,0);
        for i in 5..length loop
          Read(f,b);
        end loop;
      when others =>
        for i in 1..length loop -- just skip the contents
          Read(f,b);
        end loop;
    end case;
    New_Line;
  end loop;
  Close(f);
end;