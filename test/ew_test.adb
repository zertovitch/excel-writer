with Excel_Out;

with Ada.Strings.Fixed;

procedure EW_Test is

  use Excel_Out;

  procedure Test_Define_Col_Width (ef : Excel_Type) is
    xl : Excel_Out_File;
  begin
    xl.Create ("With def col width [" & ef'Image & "].xls", ef);
    xl.Write_Default_Column_Width (20);
    xl.Write_Column_Width (1, 5);
    xl.Write_Column_Width (5, 10, 4);
    xl.Put ("A");
    xl.Put ("B");
    xl.Close;
    --
    xl.Create ("Without def col width [" & ef'Image & "].xls", ef);
    xl.Write_Column_Width (1, 5);
    xl.Write_Column_Width (5, 10, 4);
    xl.Put ("A");
    xl.Put ("B");
    xl.Close;
  end Test_Define_Col_Width;

  --  Test automatic choice for integer output
  --
  procedure Test_Integer (ef : Excel_Type) is
    use Ada.Strings, Ada.Strings.Fixed;
    xl : Excel_Out_File;
  begin
    xl.Create ("Integer [" & ef'Image & "].xls", ef);
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
    xl.Put_Line
      ("Formulas for checking (all results should be 0). " &
       " NB: the formulas are written as text." &
       " Edit cells to convert them into real formulas.");
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
  end Test_Integer;

  procedure Test_Formulas (ef : Excel_Type) is
    xl : Excel_Out_File;
    procedure Show (formula : String) is
    begin
      xl.Put (formula);
      xl.Put_Line ("Formula: " & formula);
    end Show;
  begin
    xl.Create ("Formulas [" & ef'Image & "].xls", ef);
    Show ("=1");
    Show ("=2");
    xl.Close;
  end Test_Formulas;

begin
  for ef in Excel_Type loop
    Test_Define_Col_Width (ef);
    Test_Integer (ef);
    Test_Formulas (ef);
  end loop;
end EW_Test;
