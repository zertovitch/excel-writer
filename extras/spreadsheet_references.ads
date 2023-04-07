--  Freeware, author: G. de Montmollin

package Spreadsheet_References is

  --  References in spreadsheets are usually
  --  encoded in one of the following ways:
  --
  --    "A1"  : column is A, B, C, ...; row is 1, 2, 3, 4, ...
  --
  --    "R1C1": 'R', the row number, 'C', the column number.

  type Reference_Style is (A1, R1C1);

  function Encode_Reference
    (row, column : Positive;
     style       : Reference_Style := A1)
  return String;

  Invalid_spreadsheet_reference : exception;

  function Decode_Row (reference : String) return Positive;

  function Decode_Column (reference : String) return Positive;

  procedure Split (reference : String; row, column : out Positive);

end Spreadsheet_References;
