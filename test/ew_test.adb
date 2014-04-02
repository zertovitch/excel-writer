with Excel_Out; use Excel_Out;

procedure EW_Test is

  -- Test automatic choice for integer output
  --
  procedure Test_General(ef: Excel_type) is
    xl: Excel_Out_File;
  begin
    Create(xl, "Integer [" & Excel_type'Image(ef) & "].xls", ef);
    for power in 0 .. 66 loop
      Put(xl, power);
      Next(xl);
      Put(xl, - (2.0 ** power));
      Put(xl,    2.0 ** power );
      Put(xl,   (2.0 ** power) - 1.0 );
      Next(xl);
      if power <= 30 then
        Put(xl, - (2 ** power));
        Put(xl,    2 ** power );
        Put(xl,   (2 ** power) - 1 );
      end if;
      New_Line(xl);
    end loop;
    Close(xl);
  end;

begin
  for ef in Excel_type loop
    Test_General(ef);
  end loop;
end;
