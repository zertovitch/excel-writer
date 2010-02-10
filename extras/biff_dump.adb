-- Dump the contents of a file in BIFF (Excel .xls) format

with Excel_Out;                         use Excel_Out;

with Ada.Command_Line;                  use Ada.Command_Line;
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

  xl: Excel_Out_File;
  fmt_ul: Format_type;

begin
  if Argument_Count = 0 then
    Open(f, In_File, "big.xls");
  else
    Open(f, In_File, Argument(1));
  end if;
  Create(xl, "$Dump$.xls");
  Define_format(xl, Default_font(xl), general, fmt_ul, border => bottom);
  --
  Put_Line(xl, "Dump of the BIFF (Excel .xls) file: " & Name(f));
  New_Line(xl);
  --
  Use_format(xl, fmt_ul);
  Put(xl, "Code");
  Put(xl, "Length");
  Put(xl, " ");
  Put_Line(xl, "Comments");
  --
  Use_format(xl, Default_format(xl));
  while not End_of_File(f) loop
    code  := in16;
    length:= in16;
    Put(xl, code, base => 16);
    Put(xl, length);
    Put(xl, "    ");
    case code is
      when 16#0009# => Put(xl, "BOF"); Put(xl, "Beginning of File (Excel 2.1, BIFF2)");
      when 16#0209# => Put(xl, "BOF"); Put(xl, "Beginning of File (Excel 3.0, BIFF3)");
      when 16#0409# => Put(xl, "BOF"); Put(xl, "Beginning of File (Excel , BIFF4)");
      when 16#0809# => Put(xl, "BOF"); Put(xl, "Beginning of File (Excel , BIFF5/8)");
      when 16#000A# => Put(xl, "EOF"); Put(xl, "End of File");
      when 16#0000# => Put(xl, "DIMENSION");
      when 16#000D# => Put(xl, "CALCMODE");
      when 16#000F# => Put(xl, "REFMODE");
      when 16#0022# => Put(xl, "DATEMODE");
      when 16#0042# => Put(xl, "CODEPAGE");
      when 16#0024# => Put(xl, "COLWIDTH");
      when 16#0055# => Put(xl, "DEFCOLWIDTH");
      when 16#0025# => Put(xl, "DEFAULTROWHEIGHT");
      when row      => Put(xl, "ROW");
      when 16#001E# => Put(xl, "FORMAT");
      when 16#001F# => Put(xl, "BUILTINFMTCOUNT");
      when 16#0031# => Put(xl, "FONT");
      when 16#0045# => Put(xl, "FONTCOLOR");
      when 16#0001# => Put(xl, "BLANK");
      when 16#0002# => Put(xl, "INTEGER");
      when 16#0003# => Put(xl, "NUMBER");
      when 16#0004# => Put(xl, "LABEL");
      when xf_2 |       -- Extended Format, BIFF2
           xf_3     =>  -- Extended Format, BIFF3
        Put(xl, "XF");
        xfs:= xfs + 1;
        Put(xl, Integer'Image(xfs));
      when 16#0019# => Put(xl, "WINDOWPROTECT");
      when 16#0040# => Put(xl, "BACKUP");
      when style    => Put(xl, "STYLE");
      when others =>   Put(xl, "- ??? -");
    end case;
    case code is
      when row =>
        Put(xl, "  row="); Put(xl, in16,0);
        Put(xl, " col1="); Put(xl, in16,0);
        Put(xl, " col2="); Put(xl, in16,0);
        Put(xl, " height="); Put(xl, in16,0);
        for i in 1..5 loop
          Read(f,b);
        end loop;
      when 1..4 =>
        Put(xl, "  row="); Put(xl, in16,0);
        Put(xl, "  col="); Put(xl, in16,0);
        for i in 5..length loop
          Read(f,b);
        end loop;
      when style => -- 5.103 STYLE p. 212
        x:= in16;
        Put(xl, "  xf="); Put(xl, x mod 16#8000#, 3);
        if x >= 16#8000# then
          Put(xl, ";  built-in style: ");
          Read(f,b);
          case b is
            when 0 => Put(xl, "Normal");
            when 3 => Put(xl, "Comma");
            when 4 => Put(xl, "Currency");
            when 5 => Put(xl, "Percent");
            when others => Put(xl, Unsigned_8'Image(b));
          end case;
          for i in 4..length loop -- skip other contents
            Read(f,b);
          end loop;
        else
          Put(xl, ";  user: " & str8);
        end if;
      when xf_2  =>
        Read(f,b);
        Put(xl, "Font #" & Unsigned_8'Image(b));
        for i in 2..length loop -- skip remaining contents
          Read(f,b);
        end loop;
      when xf_3 =>
        Read(f,b);
        Put(xl, "Font #" & Unsigned_8'Image(b));
        for i in 2..length loop -- skip remaining contents
          Read(f,b);
        end loop;
      when others =>
        Put(xl, " skipping contents");
        for i in 1..length loop -- just skip the contents
          Read(f,b);
        end loop;
    end case;
    New_Line(xl);
  end loop;
  Close(f);
  Close(xl);
exception
  when others =>
    if Is_Open(f) then
      Close(f);
    end if;
    if Is_Open(xl) then
      Close(xl);
    end if;
    raise;
end;