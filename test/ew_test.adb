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
      ("Formulas for checking (all results should be 0).");
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

    function Is_Int (s : String) return Boolean is
    begin
      declare
        test : constant Integer := Integer'Value (s);
        pragma Unreferenced (test);
      begin
        return True;
      end;
    exception
      when others =>
        return False;
    end Is_Int;

    procedure Show (formula, expected_result : String) is
    begin
      xl.Put (formula);
      xl.Put (' ' & formula);
      if Is_Int (expected_result) then
        xl.Put (Integer'Value (expected_result));
      else
        xl.Put (expected_result);
      end if;
      --  !!  Add: IF(a12=c12;"";"WRONG")
      xl.New_Line;
    end Show;

  begin
    xl.Create ("Formulas [" & ef'Image & "].xls", ef);
    xl.Write_Column_Width (1, 3, 50);
    xl.Put ("Cell with formula");
    xl.Put ("How the formula should look like");
    xl.Put_Line ("How the result should look like");
    xl.Freeze_Top_Row;
    Show ("=2*4+5",          "13");
    Show ("=2+4*5",          "22");
    Show ("=(2+4)*5",        "30");
    Show ("=2^3",            "8");
    Show ("=A2",             "13");
    Show ("=A2+A3+A4+A5+A6", "86");
    Show ("=BC123",   "0");
    Show ("=BC$123",  "0");
    Show ("=$BC123",  "0");
    Show ("=$BC$123", "0");
    Show ("=""some""&"" string""", "some string");
    Show
      ("=A2&"", ""&A3&"", ""&A4&"", ""&A5&"", ""&A6&"", ""&A12",
       "13, 22, 30, 8, 13, some string");
    xl.Close;
  end Test_Formulas;

begin
  for ef in Excel_Type loop
    Test_Define_Col_Width (ef);
    Test_Integer (ef);
    Test_Formulas (ef);
  end loop;
end EW_Test;
