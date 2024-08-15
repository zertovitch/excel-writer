with Ada.Characters.Handling;
with Ada.Containers.Vectors;
with Ada.IO_Exceptions;
with Ada.Strings.Unbounded;
with Ada.Text_IO;
with Interfaces;

package body Excel_Out.Formulas is

  function Parse_Formula (text : String) return Byte_Buffer is

    use Ada.Strings.Unbounded;
    use Interfaces;

    --  Ripped from HAC_Sys.Defs:

    type KeyWSymbol is
      (IntCon,
       FloatCon,
       --  CharCon,
       StrCon,
       --
       Plus,     --  +
       Minus,    --  -
       Times,    --  *
       Divide,   --  /
       Power,    --  ^
       --
       EQL,      --  =
       NEQ,      --  /=
       GTR,      --  >
       GEQ,      --  >=
       LSS,      --  <
       LEQ,      --  <=
       --
       LParent,
       RParent,
       LBrack,
       RBrack,
       Apostrophe,
       Comma,
       Semicolon,
       Period,
       Range_Double_Dot_Symbol,  --  ".." compound delimiter (RM 2.2)
       Colon,
       Alt,
       Finger,
       Becomes,
       IDent,
       Dummy_Symbol,       --  Symbol that is never scanned.
       Ampersand_Symbol,
       NULL_Symbol);

       pragma Unreferenced (Apostrophe);

    --  Ripped from HAC_Sys.Scanner:

    type SSTBzz is array (Character'(' ') .. '^') of KeyWSymbol;

    Special_Symbols : constant SSTBzz :=
     ('+'    => Plus,
      '-'    => Minus,
      '*'    => Times,
      '/'    => Divide,
      '('    => LParent,
      ')'    => RParent,
      '['    => LBrack,
      ']'    => RBrack,
      ','    => Comma,
      ';'    => Semicolon,
      '&'    => Ampersand_Symbol,
      '^'    => Power,
      others => NULL_Symbol);

    type CHTP is (Letter, Number, Special, Illegal);

    type Set_of_CHTP is array (CHTP) of Boolean;

    special_or_illegal : constant Set_of_CHTP :=
     (Letter  |  Number  => False,
      Special | Illegal  => True);

    c128 : constant Character := Character'Val (128);

    Character_Types : constant array (Character) of CHTP :=
         ('A' .. 'Z' | 'a' .. 'z' => Letter,
          '0' .. '9' => Number,
          '#' |
          '+' | '-' | '*' | '/' |
          '(' | ')' |
          '[' | ']' |
          '&' |
          '=' |
          ' ' |
          ',' |
          '.' |
          ''' |
          ':' |
          '_' |
          ';' |
          '|' |
          '<' |
          '>' |
          '"' |
          '^' |
          '$' => Special,
          c128 => Special,
          others => Illegal);

    package Byte_Vectors is new Ada.Containers.Vectors (Positive, Unsigned_8);

    type Compiler_Data is record
      text : Unbounded_String;
      pos  : Natural := 0;
      c : Character;
      Sy, prev_sy : KeyWSymbol;
      RNum : Long_Float;
      INum : Integer;
      err_msg : Unbounded_String;
      cur_str : Unbounded_String;
      Id_with_case : Unbounded_String;
      Id     : Unbounded_String;
      output : Byte_Vectors.Vector;
    end record;

    procedure Error (CD : in out Compiler_Data; message : String) is
    begin
      CD.err_msg := To_Unbounded_String (message);
    end Error;

    procedure NextCh (CD : in out Compiler_Data) is
    begin
      if CD.pos >= Length (CD.text) then
        raise Ada.IO_Exceptions.End_Error;
      end if;
      CD.pos := CD.pos + 1;
      CD.c := Element (CD.text, CD.pos);
      if Character'Pos (CD.c) < Character'Pos (' ') then
        Error (CD, "control character");
      end if;
    end NextCh;

    procedure Skip_Blanks (CD : in out Compiler_Data) is
    begin
      while CD.pos < Length (CD.text) and then CD.c = ' ' loop
        NextCh (CD);
      end loop;
    end Skip_Blanks;

    procedure InSymbol (CD : in out Compiler_Data) is
      K, e : Integer;

      integer_digits_max : constant := 18;       --  Maximum digits for an integer literal
      EMax : constant :=  308;
      EMin : constant := -308;

      identifier_length_max : constant := 255;

      procedure Read_Scale (allow_minus : Boolean) is
        S, Sign : Integer;
        digit_count : Natural := 0;
      begin
        NextCh (CD);
        Sign := 1;
        S    := 0;
        case CD.c is
          when '+' =>
            NextCh (CD);
          when '-' =>
            NextCh (CD);
            if allow_minus then
              Sign := -1;
            else
              Error
                (CD, "negative_exponent_for_integer_literal: " &
                 CD.INum'Image & ".0e- ...");
            end if;
          when others =>
            null;
        end case;
        if CD.c not in '0' .. '9' then
          Error
            (CD,
             "illegal_character_in_number; expected digit after 'E'");
        else
          loop
            if digit_count = integer_digits_max then
              Error (CD, "integer_literal_too_large");
            elsif digit_count > integer_digits_max then
              null;  --  The insult was already issued on digit_count = integer_digits_max...
            else
              S := S * 10 + Character'Pos (CD.c) - Character'Pos ('0');
            end if;
            digit_count := digit_count + 1;
            NextCh (CD);
            exit when CD.c not in '0' .. '9';
          end loop;
        end if;
        e := S * Sign + e;
      end Read_Scale;

      procedure Adjust_Scale is
        S    : Integer;
        D, T : Long_Float;
      begin
        if K + e > EMax then
          Error
            (CD, "exponent_too_large" &
             Integer'Image (K) & " +" &
             Integer'Image (e) & " =" &
             Integer'Image (K + e) & " > Max =" &
             Integer'Image (EMax));
        elsif K + e < EMin then
          CD.RNum := 0.0;
        else
          S := abs e;
          T := 1.0;
          D := 10.0;
          loop
            while S rem 2 = 0 loop
              S := S / 2;
              D := D ** 2;
            end loop;
            S := S - 1;
            T := D * T;
            exit when S = 0;
          end loop;
          CD.RNum := (if e >= 0 then CD.RNum * T else CD.RNum / T);
        end if;
      end Adjust_Scale;

      procedure Skip_possible_underscore is
      begin
        if CD.c = '_' then
          NextCh (CD);
          if CD.c = '_' then
            Error
              (CD,
               "double_underline_not_permitted");
          elsif Character_Types (CD.c) /= Number then
            Error (CD, "digit_expected");
          end if;
        end if;
      end Skip_possible_underscore;

      procedure Read_Decimal_Float is
      begin
        --  Floating-point number 123.456
        --  Cursor is here -----------^
        if CD.c = '.' then
          --  After all, this is not a number with a decimal point,
          --  but a double dot, like 123..456.
          CD.c := c128;
          return;
        end if;
        --  Read decimal part.
        CD.Sy := FloatCon;
        CD.RNum := Long_Float (CD.INum);
        e := 0;
        while Character_Types (CD.c) = Number loop
          e := e - 1;
          CD.RNum :=
            10.0 * CD.RNum +
              Long_Float (Character'Pos (CD.c) - Character'Pos ('0'));
          NextCh (CD);
          Skip_possible_underscore;
        end loop;
        if e = 0 then
          Error (CD, "illegal_character_in_number; expected digit after '.'");
        end if;
        if CD.c in 'E' | 'e' then
          Read_Scale (allow_minus => True);
        end if;
        if e /= 0 then
          Adjust_Scale;
        end if;
      end Read_Decimal_Float;

      procedure Scan_Number (skip_leading_integer : Boolean) is
      begin
        K       := 0;
        CD.INum := 0;
        CD.Sy   := IntCon;
        if skip_leading_integer then
          --  For literals like ".123".
          Read_Decimal_Float;
        else
          --  Scan the integer part of the number.
          loop
            if K = integer_digits_max then
              Error (CD, "integer_literal_too_large");
            elsif K > integer_digits_max then
              null;  --  The insult was already issued on K = integer_digits_max...
            else
              CD.INum := CD.INum * 10 + (Character'Pos (CD.c) - Character'Pos ('0'));
            end if;
            K := K + 1;
            NextCh (CD);
            Skip_possible_underscore;
            exit when Character_Types (CD.c) /= Number;
          end loop;
          --  Integer part is read (CD.INum).
          case CD.c is
            when '.' =>
              NextCh (CD);
              Read_Decimal_Float;
            when 'E' | 'e' =>
              --  Integer with exponent: 123e4.
              e := 0;
              Read_Scale (allow_minus => False);
              --  NB: a negative exponent issues an error, then e is set to 0.
              if e > 0 then
                if K + e > integer_digits_max then
                  Error
                    (CD, "exponent_too_large" &
                     Integer'Image (K) & " +" &
                     Integer'Image (e) & " =" &
                     Integer'Image (K + e) & " > Max =" &
                     integer_digits_max'Image);
                else
                  CD.INum := CD.INum * 10 ** e;
                end if;
              end if;
            when others =>
              null;  --  Number was an integer in base 10.
          end case;
        end if;
        if Character_Types (CD.c) = Letter then
          Error (CD, "space_missing_after_number");
        end if;
      end Scan_Number;

      procedure Scan_String_Literal is
      begin
        CD.cur_str := Null_Unbounded_String;
        loop
          NextCh (CD);
          if CD.c = '"' then
            NextCh (CD);
            if CD.c /= '"' then  --  The ""x case
              exit;
            end if;
          end if;
          CD.cur_str := CD.cur_str & CD.c;
        end loop;
        CD.Sy := StrCon;
      end Scan_String_Literal;

      function To_Upper (Item : Unbounded_String) return Unbounded_String is
      begin
        return To_Unbounded_String (Ada.Characters.Handling.To_Upper (To_String (Item)));
      end To_Upper;

      exit_big_loop : Boolean;

    begin  --  InSymbol
      CD.prev_sy     := CD.Sy;

      Big_loop :
      loop
        Small_loop :
        loop
          Skip_Blanks (CD);

          exit Small_loop when Character_Types (CD.c) /= Illegal;
          Error (CD, "illegal_character [1]: [" & CD.c & ']');
          NextCh (CD);
        end loop Small_loop;

        exit_big_loop := True;
        case CD.c is
          when 'A' .. 'Z' |   --  Identifier or keyword
               'a' .. 'z' | '$' =>
            K := 0;
            CD.Id_with_case := Null_Unbounded_String;
            loop
              if K < identifier_length_max then
                K := K + 1;
                CD.Id_with_case := CD.Id_with_case & CD.c;
                if K > 1 and then Slice (CD.Id_with_case, K - 1, K) = "__" then
                  Error (CD, "double_underline_not_permitted");
                end if;
              else
                Error (CD, "identifier_too_long");
              end if;
              NextCh (CD);
              exit when CD.c /= '_'
                and then CD.c /= '$'
                and then special_or_illegal (Character_Types (CD.c));
            end loop;
            if K > 0 and then Element (CD.Id_with_case, K) = '_' then
              Error (CD, "identifier_cannot_end_with_underline");
            end if;
            CD.Id := To_Upper (CD.Id_with_case);
            --
            CD.Sy := IDent;

          when '0' .. '9' => Scan_Number (skip_leading_integer => False);
          when '"'        => Scan_String_Literal;

          when ':' =>
            NextCh (CD);
            if CD.c = '=' then
              CD.Sy := Becomes;
              NextCh (CD);
            else
              CD.Sy := Colon;
            end if;

          when '<' =>
            NextCh (CD);
            if CD.c = '=' then
              CD.Sy := LEQ;
              NextCh (CD);
            else
              CD.Sy := LSS;
            end if;

          when '>' =>
            NextCh (CD);
            if CD.c = '=' then
              CD.Sy := GEQ;
              NextCh (CD);
            else
              CD.Sy := GTR;
            end if;

          when '/' =>
            NextCh (CD);
            if CD.c = '=' then
              CD.Sy := NEQ;
              NextCh (CD);
            else
              CD.Sy := Divide;
            end if;

          when '.' =>
            NextCh (CD);
            case CD.c is
              when '.' =>
                CD.Sy := Range_Double_Dot_Symbol;
                NextCh (CD);
              when '0' .. '9' =>
                Scan_Number (skip_leading_integer => True);
              when others =>
                CD.Sy := Period;
            end case;

          when c128 =>  --  Hathorn
            CD.Sy := Range_Double_Dot_Symbol;
            NextCh (CD);

          when '-' =>
            NextCh (CD);
            CD.Sy := Minus;

          when '=' =>
            NextCh (CD);
            if CD.c = '>' then
              CD.Sy := Finger;
              NextCh (CD);
            else
              CD.Sy := EQL;
            end if;

          when '|' =>
            CD.Sy := Alt;
            NextCh (CD);

          when '+' | '*' | '(' | ')' | ',' | '[' | ']' | ';' | '&' | '^' =>
            CD.Sy := Special_Symbols (CD.c);
            NextCh (CD);

          when '!' | '@' | '\' | '_' | '?' | '%' | '#' =>
            Error (CD, "illegal_character");
            NextCh (CD);
            exit_big_loop := False;

          when Character'Val (0) .. ' ' =>
            null;
          when others =>
            null;

        end case;  --  CD.SD.CH
        exit Big_loop when exit_big_loop;
      end loop Big_loop;

    end InSymbol;

    subtype Plus_Minus is KeyWSymbol range Plus .. Minus;

    type Symset is array (KeyWSymbol) of Boolean;

    binary_adding_operator : constant Symset :=       --  RM 4.5 (4)
      (Plus | Minus | Ampersand_Symbol => True,
       others => False);

    multiplying_operator : constant Symset :=         --  RM 4.5 (6)
      (Times | Divide => True,
       others => False);

    --  3.4.1 Unary Operator Tokens, p.40
    tUplus  : constant := 16#12#;
    tUminus : constant := 16#13#;

    --  3.4.2 Binary Operator Tokens, p.40
    tAdd    : constant := 16#03#;
    tSub    : constant := 16#04#;
    tMul    : constant := 16#05#;
    tDiv    : constant := 16#06#;
    tPower  : constant := 16#07#;
    tConcat : constant := 16#08#;

    --  3.4.4 Constant Operand Tokens, p.41
    tStr : constant := 16#17#;  --  3.8.2
    tInt : constant := 16#1E#;  --  3.8.5

    --  3.9 Operand Tokens, p.54
    tRefV : constant := 16#44#;  --  3.9.2

    --  3.10 Control Tokens, p.65
    tParen : constant := 16#15#;  --  3.10.3

    --  --  Example of page 30 (2*4+5):
    --  test_data : constant Byte_Buffer :=
    --    tInt & Intel_16 (2) &
    --    tInt & Intel_16 (4) &
    --    tMul &
    --    tInt & Intel_16 (5) &
    --    tAdd;

    procedure Emit (CD : in out Compiler_Data; code : Unsigned_8) is
    begin
      CD.output.Append (code);
    end Emit;

    procedure Emit (CD : in out Compiler_Data; codes : Byte_Buffer) is
    begin
      for code of codes loop
        CD.output.Append (code);
      end loop;
    end Emit;

    type Typen is (Undefined, Ints, Floats, String_Literals);

    subtype Numeric_Typ is Typen range Ints .. Floats;

    procedure Ident_or_Cell_Reference
      (CD : in out Compiler_Data;
       X  :    out Typen)
    is
      col : Natural := 0;
      is_col_abs  : Boolean := False;
      row : Natural := 0;
      is_row_abs  : Boolean := False;
      pos : Positive;
      id : constant String := To_String (CD.Id);
    begin
      X := Undefined;
      pos := id'First;
      if pos <= id'Last and then id (pos) = '$' then
        is_col_abs := True;
        pos := pos + 1;
      end if;
      while pos <= id'Last and then id (pos) in 'A' .. 'Z' loop
        col := col * 26 + Character'Pos (id (pos)) - Character'Pos ('A') + 1;
        pos := pos + 1;
      end loop;
      if pos <= id'Last and then id (pos) = '$' then
        is_row_abs := True;
        pos := pos + 1;
      end if;
      while pos <= id'Last and then id (pos) in '0' .. '9' loop
        row := row * 10 + Character'Pos (id (pos)) - Character'Pos ('0');
        pos := pos + 1;
      end loop;
      if row > 0 and col > 0 and pos = id'Last + 1 then
        Emit (CD, tRefV);
        --  3.3.3 Cell Addresses in BIFF2-BIFF5, p.38
        Emit
          (CD,
           Intel_16
             (Unsigned_16 (row - 1) +
              16#4000# * Boolean'Pos (not is_col_abs) +
              16#8000# * Boolean'Pos (not is_row_abs)));
        Emit (CD, Unsigned_8 (col - 1));
      else
        null;  --  !!  Normal identifier !!
      end if;
    end Ident_or_Cell_Reference;

    --  Ripped from HAC_Sys.Parser.Expressions

    procedure Simple_Expression
      (CD    : in out Compiler_Data;
       X     :    out Typen)
    is

      procedure Term (X : out Typen) is

        procedure Factor (X : out Typen) is

          procedure Primary (X : out Typen) is
          begin
            X := Undefined;
            case CD.Sy is
              when StrCon =>
                X := String_Literals;
                Emit (CD, tStr);
                --  2.5.2 Byte Strings (BIFF2-BIFF5), p.17:
                Emit (CD, Unsigned_8 (Length (CD.cur_str)));
                for i in 1 .. Length (CD.cur_str) loop
                  Emit (CD, Character'Pos (Element (CD.cur_str, i)));
                end loop;
                InSymbol (CD);
              when IDent =>
                InSymbol (CD);
                Ident_or_Cell_Reference (CD, X);
              when IntCon =>  --  Literal integer or float.
                X := Ints;
                InSymbol (CD);
                Emit (CD, tInt);
                Emit (CD, Intel_16 (Unsigned_16 (CD.INum)));
              when FloatCon =>  --  Literal float.
                InSymbol (CD);
                --  !!  Process
              when LParent =>
                --  '(' : what is inside the parentheses is an
                --        expression of the lowest level.
                InSymbol (CD);
                Simple_Expression (CD, X);
                if CD.Sy = Comma then
                  Error (CD, "No aggregates");
                end if;
                if CD.Sy = RParent then
                  InSymbol (CD);
                else
                  Error (CD, "need ')'");
                end if;
                Emit (CD, tParen);
              when others =>
                null;
            end case;
          end Primary;

          Y : Typen;

        begin  --  Factor
          Primary (X);
          if CD.Sy = Power then
            InSymbol (CD);
            Primary (Y);
            Emit (CD, tPower);
          end if;
        end Factor;

        Mult_OP : KeyWSymbol;
        Y       : Typen;
      begin  --  Term
        Factor (X);
        --
        --  We collect here possible factors: a {* b}
        --
        while multiplying_operator (CD.Sy) loop
          Mult_OP := CD.Sy;
          InSymbol (CD);
          Factor (Y);
          if X in Numeric_Typ and then Y in Numeric_Typ then
            Emit (CD, (if Mult_OP = Times then tMul else tDiv));
          elsif Y not in Numeric_Typ then
            --  N * (something non-numeric)
            Error (CD, "right operand is not numeric");
          else
            Error (CD, "left operand is not numeric");
          end if;
        end loop;
      end Term;

      additive_operator : KeyWSymbol;
      y                 : Typen;

    begin  --  Simple_Expression
      if CD.Sy in Plus_Minus then
        --
        --  Unary + , -      RM 4.5 (5), 4.4 (4)
        --
        additive_operator := CD.Sy;
        InSymbol (CD);
        Term (X);
        --  At this point we have consumed "+X" or "-X".
        Emit
          (CD,
           (case Plus_Minus (additive_operator) is
            when Plus => tUplus, when Minus => tUminus));
      else
        Term (X);
      end if;
      --
      --  We collect here possible terms: a {+ b}
      --
      while binary_adding_operator (CD.Sy) loop
        additive_operator := CD.Sy;
        InSymbol (CD);
        Term (y);
        Emit
          (CD,
           (case additive_operator is
            when Plus             => tAdd,
            when Minus            => tSub,
            when Ampersand_Symbol => tConcat,
            when others           => tAdd));  --  Dummy
      end loop;
    end Simple_Expression;

    procedure Parse_Init (CD : in out Compiler_Data) is
    begin
      CD.c := ' ';
      CD.pos := 1;
      CD.Sy := Dummy_Symbol;
      InSymbol (CD);
    end Parse_Init;

    X : Typen;
    CD : Compiler_Data;

    trace : constant Boolean := False;

  begin
    CD.text := To_Unbounded_String (' ' & text & ' ');
    Parse_Init (CD);
    Simple_Expression (CD, X);
    if trace then
      Ada.Text_IO.Put_Line (text);
      for elem of CD.output loop
        Ada.Text_IO.Put_Line (elem'Image);
      end loop;
      if CD.err_msg /= "" then
        Ada.Text_IO.Put_Line (To_String (CD.err_msg));
      end if;
    end if;
    declare
      buf : Byte_Buffer (1 .. Integer (CD.output.Length));
    begin
      for i in 1 .. Integer (CD.output.Length) loop
        buf (i) := CD.output (i);
      end loop;
      return buf;
    end;
  end Parse_Formula;

end Excel_Out.Formulas;
