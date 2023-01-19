Excel Writer 
============

The Excel Writer consists of an Ada package, Excel_Out,
which produces Excel files - as physical files, or as
other types of data streams.

The creation of an Excel file is as simple as this
small procedure:

  with Excel_Out;
   
  procedure Small_demo is 
    xl : Excel_Out.Excel_Out_File;
  begin 
    xl.Create ("Small.xls");
    xl.Put_Line ("Hello world !");
    xl.Close;
  end;


License
=======
Excel Writer is licensed under the MIT License.


====

Full description in: excel_writer.txt
