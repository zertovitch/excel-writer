with Excel_Out;

with Ada.Strings.Fixed;

procedure EW_Test is

  use Excel_Out;

  procedure Test_def_col_width (ef : Excel_type) is
    xl : Excel_Out_File;
  begin
    xl.Create ("With def col width [" & Excel_type'Image (ef) & "].xls", ef);
    xl.Write_default_column_width (20);
    xl.Write_column_width (1, 5);
    xl.Write_column_width (5, 10, 4);
    xl.Put ("A");
    xl.Put ("B");
    xl.Close;
    --
    xl.Create ("Without def col width [" & Excel_type'Image (ef) & "].xls", ef);
    xl.Write_column_width (1, 5);
    xl.Write_column_width (5, 10, 4);
    xl.Put ("A");
    xl.Put ("B");
    xl.Close;
  end Test_def_col_width;

  --  Test automatic choice for integer output
  --
  procedure Test_General (ef : Excel_type) is
    use Ada.Strings, Ada.Strings.Fixed;
    xl : Excel_Out_File;
  begin
    xl.Create ("Integer [" & Excel_type'Image (ef) & "].xls", ef);
    xl.Freeze_Top_Row;
    xl.Put ("x");
    xl.Next;
    xl.Put  ("2.0**x");
    xl.Put ("-2.0**x");
    xl.Put  ("2.0**x - 1");
    xl.Next;
    xl.Put  ("2**x");
    xl.Put ("-2**x");
    xl.Put  ("2**x - 1");
    xl.Next;
    xl.Put_Line ("Formulas for checking (all results should be 0)");
    for power in 0 .. 66 loop
      xl.Put (power);
      xl.Next;
      xl.Put   (2.0 ** power);
      xl.Put (-(2.0 ** power));
      xl.Put  ((2.0 ** power) - 1.0);
      xl.Next;
      if power <= 30 then
        xl.Put   (2 ** power);
        xl.Put (-(2 ** power));
        xl.Put  ((2 ** power) - 1);
      else
        xl.Next (3);
      end if;
      xl.Next;
      declare
        row : constant String := Trim (Integer'Image (power + 2), Both);
      begin
        xl.Put ("= (2^A" & row & ")     - C" & row);
        xl.Put ("=-(2^A" & row & ")     - D" & row);
        xl.Put ("= (2^A" & row & ") - 1 - E" & row);
        xl.Next;
        if power <= 30 then
          xl.Put ("= (2^A" & row & ")     - G" & row);
          xl.Put ("=-(2^A" & row & ")     - H" & row);
          xl.Put ("= (2^A" & row & ") - 1 - I" & row);
        end if;
      end;
      xl.New_Line;
    end loop;
    xl.Close;
  end Test_General;

begin
  for ef in Excel_type loop
    Test_def_col_width (ef);
    Test_General (ef);
  end loop;
end EW_Test;
