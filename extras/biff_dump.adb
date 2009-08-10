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

  code, length, x: Integer;

  function in16 return Integer is
    b1,b2: Unsigned_8;
  begin
    Read(f,b1);
    Read(f,b2);
    return Integer(b1) + Integer(b2) * 256;
  end in16;

  function str8 return String is
    b: Unsigned_8;
  begin
    Read(f,b);
    declare
      r: String(1..Integer(b));
    begin
      for i in r'Range loop
        Read(f,b);
        r(i):= character'Val(b);
      end loop;
      return r;
    end;
  end str8;

  row  : constant:= 16#0008#;
  style: constant:= 16#0293#;
  xf_2 : constant:= 16#0043#;
  xf_3 : constant:= 16#0243#;

  b: Unsigned_8;
  xfs: Natural:= 0;

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
      when 16#0009# => Put("BOF - Beginning of File (Excel 2.1, BIFF2)");
      when 16#0209# => Put("BOF - Beginning of File (Excel 3.0, BIFF3)");
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
      when xf_2 |       -- Extended Format, BIFF2
           xf_3     =>  -- Extended Format, BIFF3
        Put("XF");
        xfs:= xfs + 1;
        Put(Integer'Image(xfs) & ", ");
      when 16#0019# => Put("WINDOWPROTECT");
      when 16#0040# => Put("BACKUP");
      when style    => Put("STYLE");
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
      when style => -- 5.103 STYLE p. 212
        x:= in16;
        Put("  xf="); Put(x mod 16#8000#, 3);
        if x >= 16#8000# then
          Put(";  built-in style: ");
          Read(f,b);
          case b is
            when 0 => Put("Normal");
            when 3 => Put("Comma");
            when 4 => Put("Currency");
            when 5 => Put("Percent");
            when others => Put(Unsigned_8'Image(b));
          end case;
          for i in 4..length loop -- skip other contents
            Read(f,b);
          end loop;
        else
          Put(";  user: " & str8);
        end if;
      when xf_2  =>
        Read(f,b);
        Put("Font #" & Unsigned_8'Image(b));
        for i in 2..length loop -- skip remaining contents
          Read(f,b);
        end loop;
      when xf_3 =>
        Read(f,b);
        Put("Font #" & Unsigned_8'Image(b));
        for i in 2..length loop -- skip remaining contents
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