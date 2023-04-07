with Ada.Text_IO;
with Spreadsheet_References;

procedure Spreadsheet_References_Demo is
  ii, jj, n0 : Positive;
  use Ada.Text_IO, Spreadsheet_References;

  function Display_both_encodings (i, j : Positive) return String is
  begin
    return Encode_Reference (i, j, A1) & " = " & Encode_Reference (i, j, R1C1);
  end Display_both_encodings;

begin
  Split ("xfd1234", ii, jj);
  Put_Line (Integer'Image (ii) & Integer'Image (jj));
  New_Line;
  for i in 1 .. 10 loop
    Put_Line (Display_both_encodings (i, 1));
  end loop;
  New_Line;
  for j in 1 .. 10 loop
    Put_Line (Display_both_encodings (1, j));
  end loop;
  New_Line;
  for i in 1 .. 256 loop
    Put_Line (Display_both_encodings (i, i));
  end loop;
  New_Line;
  --  Excel 2007 and later has 16384 = 2**16 instead of 256 = 2**8 columns
  for i in 700 .. 710 loop -- ZZ=702, AAA=703
    Put_Line (Display_both_encodings (i, i));
  end loop;
  New_Line;
  for i in 16382 .. 16386 loop -- XFD=16384
    Put_Line (Display_both_encodings (i, i));
  end loop;
  New_Line;
  for n in 1 .. 5 loop
    n0 := (26**n - 1) * 26 / 25;
    for i in n0 - 2 .. n0 + 2 loop
      Put_Line (Display_both_encodings (i, i));
    end loop;
    New_Line;
  end loop;
  --
  --  Consistency check on different rows, columns and on both styles
  --
  for i in 1 .. 123 loop
    for j in 1 .. 16389 loop
      for style in Reference_Style loop
        Split (Encode_Reference (i, j, style), ii, jj);
        if i /= ii or j /= jj then
          raise Program_Error;
        end if;
      end loop;
    end loop;
  end loop;
end Spreadsheet_References_Demo;
