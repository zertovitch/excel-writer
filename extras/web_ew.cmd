echo with Excel_Out; use Excel_Out;>small_demo.adb
echo.>>small_demo.adb
echo procedure Small_demo is>>small_demo.adb
echo   xl: Excel_Out_File;>>small_demo.adb
echo begin>>small_demo.adb
echo   xl.Create("Small.xls");>>small_demo.adb
echo   xl.Put_Line("This is a small demo for Excel_Out");>>small_demo.adb
echo   for row in 3 .. 8 loop>>small_demo.adb
echo     for column in 1 .. 8 loop>>small_demo.adb
echo       xl.Write(row, column, row * 1000 + column);>>small_demo.adb
echo     end loop;>>small_demo.adb
echo   end loop;>>small_demo.adb
echo   xl.Close;>>small_demo.adb
echo end Small_demo;>>small_demo.adb

rem Call GNATMake without project file: we want the .ali here.

gnatmake ..\excel_out_demo.adb -I..
gnatmake small_demo.adb -I..

rem Small_Demo without local references
perl ew_html.pl small_demo -d -I.. -oew_html
perl ew_html.pl excel_out_demo excel_out.ads excel_out.adb -I.. -f -d -oew_html

del *.ali
del *.o
