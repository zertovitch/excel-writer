with Excel_Out; use Excel_Out;

procedure EW_Test is

  -- Issue: default format in BIFF3 mode is not General
  --
  procedure Test_General(ef: Excel_type) is
    xl: Excel_Out_File;
  begin
    xl.Create("General numeric format [" & Excel_type'Image(ef) & "].xls", ef);
    for row in 3 .. 8 loop
      for column in 1 .. 8 loop
        xl.Write(row, column, row * 1000 + column);
      end loop;
    end loop;
    xl.Close;
  end;

  -- Issue: Write_row_height bad display on MS Excel if height > 0 ; LibreOffice OK
  -- BIFF2 and BIFF3 affected
  --
  procedure Test_Row(ef: Excel_type) is
    xl: Excel_Out_File;
  begin
    xl.Create("Row height [" & Excel_type'Image(ef) & "].xls", ef);
    xl.Write_row_height(1, 33);
    xl.Put("A");
    xl.Close;
  end;

begin
  for ef in Excel_type loop
    Test_General(ef);
    Test_Row(ef);
  end loop;
end;
