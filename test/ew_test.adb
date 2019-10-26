with Excel_Out; use Excel_Out;

with Ada.Strings.Fixed;                 use Ada.Strings, Ada.Strings.Fixed;

procedure EW_Test is

  procedure Test_def_col_width (ef : Excel_type) is
    xl : Excel_Out_File;
  begin
    Create (xl, "With def col width [" & Excel_type'Image (ef) & "].xls", ef);
    Write_default_column_width (xl, 20);
    Write_column_width (xl, 1, 5);
    Write_column_width (xl, 5, 10, 4);
    Put (xl, "A");
    Put (xl, "B");
    Close (xl);
    --
    Create (xl, "Without def col width [" & Excel_type'Image (ef) & "].xls", ef);
    Write_column_width (xl, 1, 5);
    Write_column_width (xl, 5, 10, 4);
    Put (xl, "A");
    Put (xl, "B");
    Close (xl);
  end Test_def_col_width;

  -- Test automatic choice for integer output
  --
  procedure Test_General (ef : Excel_type) is
    xl : Excel_Out_File;
  begin
    Create (xl, "Integer [" & Excel_type'Image (ef) & "].xls", ef);
    Freeze_Top_Row (xl);
    Put (xl, "x");
    Next (xl);
    Put (xl,  "2.0**x");
    Put (xl, "-2.0**x");
    Put (xl,  "2.0**x - 1");
    Next (xl);
    Put (xl,  "2**x");
    Put (xl, "-2**x");
    Put (xl,  "2**x - 1");
    Next (xl);
    Put_Line (xl, "Formulas for checking (all results should be 0)");
    for power in 0 .. 66 loop
      Put (xl, power);
      Next (xl);
      Put (xl,    2.0 ** power);
      Put (xl, -(2.0 ** power));
      Put (xl,   (2.0 ** power) - 1.0);
      Next (xl);
      if power <= 30 then
        Put (xl,    2 ** power);
        Put (xl, -(2 ** power));
        Put (xl,   (2 ** power) - 1);
      else
        Next (xl, 3);
      end if;
      Next (xl);
      declare
        row : constant String := Trim (Integer'Image (power + 2), Both);
      begin
        Put (xl, "= (2^A" & row & ")     - C" & row);
        Put (xl, "=-(2^A" & row & ")     - D" & row);
        Put (xl, "= (2^A" & row & ") - 1 - E" & row);
        Next (xl);
        if power <= 30 then
          Put (xl, "= (2^A" & row & ")     - G" & row);
          Put (xl, "=-(2^A" & row & ")     - H" & row);
          Put (xl, "= (2^A" & row & ") - 1 - I" & row);
        end if;
      end;
      New_Line (xl);
    end loop;
    Close (xl);
  end;

begin
  for ef in Excel_type loop
    Test_def_col_width (ef);
    Test_General (ef);
  end loop;
end;
