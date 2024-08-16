@echo off

gprbuild -p -P excel_out -XExcel_Build_Mode=Fast

if %errorlevel% == 9009 goto error

echo Press Return
pause
goto :eof


:error

echo.
echo The GNAT Ada compiler was not found in the PATH!
echo.
echo Check https://www.adacore.com/download for GNAT
echo or https://alire.ada.dev/ for ALIRE.
echo The Excel project also is available as an ALIRE crate.
echo.
echo Press Return
pause
