echo with Excel_Out;>small_demo.adb
echo.>>small_demo.adb
echo procedure Small_Demo is>>small_demo.adb
echo   xl: Excel_Out.Excel_Out_File;>>small_demo.adb
echo begin>>small_demo.adb
echo   xl.Create ("small.xls");>>small_demo.adb
echo   xl.Put_Line ("This is a small demo for Excel_Out");>>small_demo.adb
echo   for row in 3 .. 8 loop>>small_demo.adb
echo     for column in 1 .. 8 loop>>small_demo.adb
echo       xl.Write (row, column, row * 1000 + column);>>small_demo.adb
echo     end loop;>>small_demo.adb
echo   end loop;>>small_demo.adb
echo   xl.Close;>>small_demo.adb
echo end Small_Demo;>>small_demo.adb


gnatmake -P ..\excel_out.gpr

rem Call GNATMake without project file: we want the .ali here.
gnatmake small_demo.adb -I.. -aO../obj_debug -j0

set params=-oew_html -b#fffcfb -iew_head.txt -jew_top.txt -kew_bottom.txt
set params=%params% -I../obj_debug -I..

rem Small_Demo without local references (TBD: -f switch for new GNATHTML)
gnathtml small_demo.adb %params%
gnathtml excel_out_demo.adb excel_out.ads excel_out.adb -f %params%

del small_demo.ali
del small_demo.o
REM del small_demo.exe
