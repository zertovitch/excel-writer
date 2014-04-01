-- Dump the contents of a file in BIFF (Excel .xls) format

with Excel_Out;                         use Excel_Out;

with Ada.Command_Line;                  use Ada.Command_Line;
with Ada.Directories;
with Ada.Sequential_IO;
with Ada.Strings.Unbounded;             use Ada.Strings.Unbounded;

with Interfaces;                        use Interfaces;

procedure BIFF_Dump is

  package BIO is new Ada.Sequential_IO(Unsigned_8);
  use BIO;

  f: BIO.File_Type;

  code, length, x: Integer;

  function in8 return Integer is
    b: Unsigned_8;
  begin
    Read(f,b);
    return Integer(b);
  end in8;

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
        r(i):= Character'Val(b);
      end loop;
      return r;
    end;
  end str8;

  function str8len16 return String is
    r: String(1..in16);
    b: Unsigned_8;
  begin
    for i in r'Range loop
      Read(f,b);
      r(i):= Character'Val(b);
    end loop;
    return r;
  end str8len16;

  row_2      : constant:= 16#0008#; -- 5.88 p.202
  row_3      : constant:= 16#0208#; -- 5.88 p.202
  style      : constant:= 16#0293#;
  xf_2       : constant:= 16#0043#;
  xf_3       : constant:= 16#0243#;
  xf_4       : constant:= 16#0443#;
  xf_5       : constant:= 16#00E0#;
  ole_2      : constant:= 16#CFD0#;
  pane       : constant:= 16#0041#;
  window1    : constant:= 16#003D#;
  window2_b2 : constant:= 16#003E#;
  window2_b3 : constant:= 16#023E#;
  hideobj    : constant:= 16#008D#;
  font2      : constant:= 16#0031#;
  font3      : constant:= 16#0231#;
  format2    : constant:= 16#001E#;
  format4    : constant:= 16#041E#;
  blank2     : constant:= 16#0001#;
  index3     : constant:= 16#020B#; -- 5.59
  integer2   : constant:= 16#0002#; -- 5.60
  number2    : constant:= 16#0003#; -- 5.71
  number3    : constant:= 16#0203#; -- 5.71
  rk         : constant:= 16#027E#; -- 5.87 RK p.201
  note       : constant:= 16#001C#; -- 5.70 NOTE p. 190
  label2     : constant:= 16#0004#;
  label3     : constant:= 16#0204#;
  labelsst   : constant:= 16#00FD#;
  formula2   : constant:= 16#0006#; -- Formula BIFF 2, 5.50 p.176
  formula4   : constant:= 16#0406#; -- Formula BIFF 4
  colwidth   : constant:= 16#0024#;
  defcolwidth: constant:= 16#0055#;
  colinfo    : constant:= 16#007D#;
  header_x   : constant:= 16#0014#; -- 5.55 p.180
  footer_x   : constant:= 16#0015#; -- 5.48 p.173
  page_setup_x : constant:= 16#00A1#; -- 5.73 p.192
  dimension_b2 : constant:= 16#0000#;
  dimension_b3 : constant:= 16#0200#;
  writeaccess  : constant:= 16#005C#; -- 5.112 WRITEACCESS
  saverecalc   : constant:= 16#005F#; -- 5.90 SAVERECALC
  guts         : constant:= 16#0080#; -- 5.53 GUTS
  sheetpr      : constant:= 16#0081#; -- 5.97 SHEETPR
  gridset      : constant:= 16#0082#; -- 5.52 GRIDSET
  hcenter      : constant:= 16#0083#; -- 5.54 HCENTER
  vcenter      : constant:= 16#0084#; -- 5.107 VCENTER
  country      : constant:= 16#008C#; -- 5.22 COUNTRY

  subtype margin is Integer range 16#26#..16#29#;

  b: Unsigned_8;
  w: Unsigned_16;
  xfs: Natural:= 0;
  fmt: Natural:= 0;
  fnt: Natural:= 0;
  biff_version: Natural:= 0;
  defaults: Boolean;

  xl: Excel_Out_File;
  fmt_ul: Format_type;

  procedure Cell_Attributes is
  begin
    Put(xl, "xf=" & Integer'Image(in8 mod 16#40#));
    Read(f,b);
    Put(xl, "num format=" & Unsigned_8'Image(b mod 16#40#));
    Put(xl, "font="       & Unsigned_8'Image(b / 16#40#));
    Read(f,b);
  end;

  name: Unbounded_String;

begin
  if Argument_Count = 0 then
    name:= To_Unbounded_String("Big [BIFF3].xls");
  else
    name:= To_Unbounded_String(Argument(1));
  end if;
  Create(xl, "_Dump of " & Ada.Directories.Simple_Name(To_String(name)) & "", BIFF2);
  -- Some page layout...
  Header(xl, "&LBiff_dump of...&R" & Ada.Directories.Simple_Name(To_String(name)));
  Footer(xl, "&L&D");
  Margins(xl, 0.7, 0.5, 1.0, 0.8);
  Print_Gridlines(xl);
  Page_Setup(
    xl,
    orientation => landscape,
    scale_or_fit => fit,
    fit_height_with_n_pages => 0
  );
  --
  Write_default_column_width(xl, 18);
  Write_column_width(xl, 1, 11);
  Write_column_width(xl, 3, 3);
  Write_column_width(xl, 4, 20);
  --
  Define_format(xl, Default_font(xl), general, fmt_ul, border => bottom);
  --
  Put_Line(xl, "Dump of the BIFF (Excel .xls) file: " & To_String(name));
  New_Line(xl);
  --
  Use_format(xl, fmt_ul);
  Put(xl, "BIFF Code");
  Put(xl, "Bytes");
  Put(xl, " ");
  Put(xl, "BIFF Topic");
  Put_Line(xl, "Comments");
  --
  Use_format(xl, Default_format(xl));
  Open(f, In_File, To_String(name));
  while not End_Of_File(f) loop
    code  := in16;
    length:= in16;
    Put(xl, code, base => 16);
    Put(xl, length);
    Put(xl, "    ");
    case code is
      --
      when 16#0009# =>
        Put(xl, "BOF");
        Put(xl, "Beginning of File (Excel 2.1, BIFF2)");
        biff_version:= 2; -- some items, like font, are reused in biff 5 but not 3,4
      when 16#0209# =>
        Put(xl, "BOF");
        Put(xl, "Beginning of File (Excel 3.0, BIFF3)");
        biff_version:= 3;
      when 16#0409# =>
        Put(xl, "BOF");
        Put(xl, "Beginning of File (Excel 4.0, BIFF4)");
        biff_version:= 4;
      when 16#0809# =>
        Put(xl, "BOF");
        Put(xl, "Beginning of File (Excel 5-95 / 97-2003, BIFF5 / 8)");
        biff_version:= 5;
      when 16#000A# => Put(xl, "EOF"); Put(xl, "End of File");
      --
      when dimension_b2 => Put(xl, "DIMENSION (BIFF2)");  -- 5.35 DIMENSION
      when dimension_b3 => Put(xl, "DIMENSION (BIFF3+)"); -- 5.35 DIMENSION
      when 16#000C# => Put(xl, "CALCCOUNT");
      when 16#000D# => Put(xl, "CALCMODE");
      when 16#000E# => Put(xl, "PRECISION");
      when 16#000F# => Put(xl, "REFMODE");
      when 16#0010# => Put(xl, "DELTA");
      when 16#0011# => Put(xl, "ITERATION");
      when 16#002A# => Put(xl, "PRINTHEADERS");
      when 16#002B# => Put(xl, "PRINTGRIDLINES");
      when page_setup_x => Put(xl, "PAGESETUP");
      when header_x => Put(xl, "HEADER");
      when footer_x => Put(xl, "FOOTER");
      when margin   => Put(xl, "MARGIN");
      when 16#0022# => Put(xl, "DATEMODE");
      when 16#0042# => Put(xl, "CODEPAGE");
      when colwidth    => Put(xl, "COLWIDTH (BIFF2)");
      when defcolwidth => Put(xl, "DEFCOLWIDTH");
      when colinfo     => Put(xl, "COLINFO (BIFF3+)"); -- 5.18
      when 16#0025#    => Put(xl, "DEFAULTROWHEIGHT (BIFF2)");
      when 16#0225#    => Put(xl, "DEFAULTROWHEIGHT (BIFF3+)");
      when row_2 | row_3 =>
        Put(xl, "ROW");
      when format2  =>
        Put(xl, "FORMAT (BIFF2-3)" & Integer'Image(fmt));
        fmt:= fmt + 1;
      when format4  =>
        Put(xl, "FORMAT (BIFF4+)"  & Integer'Image(fmt)); -- 5.49
        fmt:= fmt + 1;
      when xf_2 |       -- Extended Format, BIFF2  -- 5.115
           xf_3 |       -- Extended Format, BIFF3
           xf_4 |       -- Extended Format, BIFF4
           xf_5     =>  -- Extended Format, BIFF5+
        Put(xl, "XF" & Integer'Image(xfs));
        xfs:= xfs + 1;
      when 16#001F# | 16#0056# =>
        Put(xl, "BUILTINFMTCOUNT");
      when font2 | font3 =>
        if fnt = 4 then
          fnt:= 5; -- Excel anomaly (p.171)
        end if;
        Put(xl, "FONT" & Integer'Image(fnt));
        -- 5.45, p.171
        fnt:= fnt + 1;
      when 16#0045#   => Put(xl, "FONTCOLOR");
      when blank2     => Put(xl, "BLANK (BIFF2)");  -- 5.7 p.137
      when 16#0201#   => Put(xl, "BLANK (BIFF3+)");
      when index3     => Put(xl, "INDEX (BIFF3+)");
      when integer2   => Put(xl, "INTEGER (BIFF2)");
      when number2    => Put(xl, "NUMBER (BIFF2)");
      when number3    => Put(xl, "NUMBER (BIFF3+)");
      when formula2   => Put(xl, "FORMULA (BIFF2)"); -- 5.50 p.176
      when formula4   => Put(xl, "FORMULA (BIFF4)");
      when rk         => Put(xl, "RK (BIFF3+)");
      when note       => Put(xl, "NOTE (Comment)"); -- 5.70 NOTE p. 190
      when label2     => Put(xl, "LABEL (BIFF2)");
      when label3     => Put(xl, "LABEL (BIFF3+)");
      when labelsst   => Put(xl, "LABELSST (BIFF8)"); -- SST = shared string table
      when 16#0019#   => Put(xl, "WINDOWPROTECT");
      when 16#0040#   => Put(xl, "BACKUP");
      when style      => Put(xl, "STYLE");            -- 5.103
      when pane       => Put(xl, "PANE");             -- 5.75 p.197
      when window1    => Put(xl, "WINDOW1");          -- 5.109
      when window2_b2 => Put(xl, "WINDOW2 (BIFF2)");  -- 5.110 p.216
      when window2_b3 => Put(xl, "WINDOW2 (BIFF3+)"); -- 5.110 p.216
      when 16#001D#   => Put(xl, "SELECTION"); -- 5.93 p.205
      when hideobj    => Put(xl, "HIDEOBJ"); -- 5.56
      when 16#4D#     => Put(xl, "PLS (Current printer blob)");
      when 16#3C#     => Put(xl, "CONTINUE (Continue last BIFF record)");
      when writeaccess => Put(xl, "WRITEACCESS"); -- 5.112 WRITEACCESS
      when saverecalc  => Put(xl, "SAVERECALC");  -- 5.90 SAVERECALC
      when gridset     => Put(xl, "GRIDSET");     -- 5.52 GRIDSET
      when guts        => Put(xl, "GUTS");        -- 5.53 GUTS
      when hcenter     => Put(xl, "HCENTER");     -- 5.107 HCENTER
      when vcenter     => Put(xl, "VCENTER");
      when sheetpr     => Put(xl, "SHEETPR");     -- 5.97 SHEETPR
      when country     => Put(xl, "COUNTRY");     -- 5.22 COUNTRY
      when others      => Put(xl, "- ??? -");
    end case;
    --
    -- Expand parameters
    --
    case code is
      when row_2 | row_3=> -- 5.88 p.202
        Put(xl, "row=" & Integer'Image(in16+1));
        Put(xl, "col1=" & Integer'Image(in16+1));
        Put(xl, "col2+1=" & Integer'Image(in16+1));
        w:= Unsigned_16(in16);
        if (w and 16#8000#) /= 0 then
          Put(xl, "default height");
        else
          Put(xl, "height=" & Float'Image(Float(w and 16#7FFF#)/20.0));
        end if;
        if biff_version = 2 then
          Put(xl, Integer'Image(in16) & " unused ?");
          Read(f,b);
          defaults:= b = 0;
          if defaults then
            Put(xl, "0: no default attributes/formats");
          else
            Put(xl, "default attributes");
          end if;
          Put(xl, "offset to contents = " & Integer'Image(in16));
          for i in 14..length loop
            Read(f,b);
          end loop;
        else
          Put(xl, "reserved1=" & Integer'Image(in16)); -- reserved1 (2 bytes): MUST be zero, and MUST be ignored.
          Put(xl, "unused1=" & Integer'Image(in16));   -- unused1 (2 bytes): Undefined and MUST be ignored.
          Put(xl, "flags=" & Integer'Image(in8));
          -- A - iOutLevel (3 bits): An unsigned integer that specifies the outline level (1) of the row.
          -- B - reserved2 (1 bit): MUST be zero, and MUST be ignored.
          -- C - fCollapsed (1 bit): A bit that specifies whether the rows that are one level of outlining deeper than the current row are included in the collapsed outline state.
          -- D - fDyZero (1 bit): A bit that specifies whether the row is hidden.
          -- E - fUnsynced (1 bit): A bit that specifies whether the row height was manually set.
          -- F - fGhostDirty (1 bit): A bit that specifies whether the row was formatted.
          Put(xl, "reserved3=" & Integer'Image(in8)); -- MUST be 1, and MUST be ignored
          Put(xl, "ixfe_val_etc=" & Integer'Image(in16));   -- ixfe_val (12 bits) and 4 bits
        end if;
      when blank2 | number2 =>
        Put(xl, "row=" & Integer'Image(in16+1));
        Put(xl, "col=" & Integer'Image(in16+1));
        Cell_Attributes;
        for i in 8..length loop
          Read(f,b);
        end loop;
      when integer2 =>
        Put(xl, "row=" & Integer'Image(in16+1));
        Put(xl, "col=" & Integer'Image(in16+1));
        Cell_Attributes;
        Put(xl, in16);
      when number3 | rk =>
        Put(xl, "row=" & Integer'Image(in16+1));
        Put(xl, "col=" & Integer'Image(in16+1));
        Put(xl, "xf="  & Integer'Image(in16));
        for i in 7..length loop
          Read(f,b);
        end loop;
      when note => -- 5.70 NOTE p. 190
        Put(xl, "row=" & Integer'Image(in16+1));
        Put(xl, "col=" & Integer'Image(in16+1));
        Put(xl, "total length=" & Integer'Image(in16+1));
        declare
          chunk: String(7..length);
        begin
          for i in chunk'Range loop
            Read(f,b);
            chunk(i):= Character'Val(b);
          end loop;
          Put(xl, chunk);
        end;
      when label2 => -- 5.63 LABEL p.187
        Put(xl, "row=" & Integer'Image(in16+1));
        Put(xl, "col=" & Integer'Image(in16+1));
        Cell_Attributes;
        Put(xl, str8);
      when label3 => -- 5.63 LABEL p.187
        Put(xl, "row=" & Integer'Image(in16+1));
        Put(xl, "col=" & Integer'Image(in16+1));
        Put(xl, "xf="  & Integer'Image(in16));
        Put(xl, str8len16);
      when labelsst => -- SST = shared string table
        Put(xl, "row=" & Integer'Image(in16+1));
        Put(xl, "col=" & Integer'Image(in16+1));
        for i in 5..length loop
          Read(f,b);
        end loop;
      when format2 =>
        Put(xl, str8);
      when font2 =>
        Put(xl, "height="  & Float'Image(Float(in16)/20.0));
        Put(xl, "options=" & Integer'Image(in16));
        if biff_version = 2 then
          declare
            font_name: constant String:= str8;
          begin
            Put(xl, font_name);
            for i in 6+font_name'Length .. length loop
              -- Excel 2002 puts garbage, sometimes...
              Read(f,b);
            end loop;
          end;
        else -- BIFF 5-8
          for i in 5..length loop -- just skip the contents
            Read(f,b);
          end loop;
        end if;
      when font3 =>
        Put(xl, "height=" & Float'Image(Float(in16)/20.0));
        Put(xl, "options=" & Integer'Image(in16));
        Put(xl, "colour="  & Integer'Image(in16));
        Put(xl, str8);
      when style => -- 5.103 STYLE p. 212
        x:= in16;
        Put(xl, "  xf=");
        Put(xl, x mod 16#2000#, 3);
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
          Read(f,b);
          Put(xl, "Level" & Unsigned_8'Image(b));
        else
          Put(xl, ";  user: " & str8);
        end if;
      when xf_2  => -- 5.115 XF - Extended Format p.219
        Read(f,b);
        Put(xl, "Using font #" & Unsigned_8'Image(b));
        Read(f,b); -- skip
        Read(f,b);
        Put(xl, "(Number) format #" & Unsigned_8'Image(b and 16#3F#));
        for i in 4..length loop -- skip remaining contents
          Read(f,b);
        end loop;
      when xf_3 =>
        Read(f,b);
        Put(xl, "Using font #" & Unsigned_8'Image(b));
        Read(f,b);
        Put(xl, "(Number) format #" & Unsigned_8'Image(b));
        Read(f,b); -- skip Protection
        Read(f,b); -- skip Used attributes
        for i in 5..length loop -- skip remaining contents
          Read(f,b);
        end loop;
      when ole_2 =>
        Put_Line(xl, "This is an OLE-OLE 2 file, eventually wrapping a BIFF one");
        Put_Line(xl, "Check: Microsoft Compound Document File Format, compdocfileformat.pdf");
        Put_Line(xl, "Aborting dump");
        Close(f);
        Close(xl);
        return;
      when colwidth =>
        Put(xl, "First Column: " & Integer'Image(in8+1));
        Put(xl, "Last Column : " & Integer'Image(in8+1));
        Put(xl, "Width: " & Float'Image(Float(in16)/256.0));
      when defcolwidth =>
        Put(xl, "Width: " & Float'Image(Float(in16)/256.0));
      when header_x | footer_x =>
        if length > 0 then
          declare
            head_foot: constant String:= str8;
          begin
            Put(xl, head_foot);
            for i in 2+head_foot'Length .. length loop
              -- garbage
              Read(f,b);
            end loop;
          end;
        end if;
      when page_setup_x =>
        Put(xl, "paper=" & Integer'Image(in16));
        Put(xl, "scaling="  & Integer'Image(in16));
        Put(xl, "start page="  & Integer'Image(in16));
        Put(xl, "fit width="  & Integer'Image(in16));
        Put(xl, "fit height="  & Integer'Image(in16));
        Put(xl, "options="  & Integer'Image(in16));
        for i in 13..length loop -- remaining contents (BIFF5+)
          Read(f,b);
        end loop;
      when dimension_b2 | dimension_b3 =>
        Put(xl, "row_min="    & Integer'Image(in16));
        Put(xl, "row_max+1="  & Integer'Image(in16));
        Put(xl, "col_min="    & Integer'Image(in16));
        Put(xl, "col_max+1="  & Integer'Image(in16));
        for i in 9..length loop -- remaining contents (BIFF3+)
          Read(f,b);
        end loop;
      when writeaccess =>
        declare
          r: constant String:= str8;
        begin
          Put(xl, "User name=" & r);
          for i in r'Length+2..length loop -- remaining characters (spaces)
            Read(f,b);
          end loop;
        end;
      when others =>
        --  if length > 0 then
        --    Put(xl, "skipping contents");
        --  end if;
        for i in 1..length loop -- just skip the contents
          if i <= 10 then
            Put(xl, in8);
          else
            Read(f,b);
          end if;
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
