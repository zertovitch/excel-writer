gnatmake ..\Excel_Out_Test.adb
gnatmake Small_demo.adb -I..

perl ew_html.pl excel_out_test excel_out.ads excel_out.adb -I.. -f -d -oew_html
perl ew_html.pl small_demo -d -I.. -o0_html

