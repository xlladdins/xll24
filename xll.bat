REM Create project zip file for template
msbuild -t:Clean
set XLL_DIR=Templates\ProjectTemplates\Visual C++
del xll.zip
REM tar -C "%XLL_DIR%" -z -c -f xll.zip .
"C:\Program Files\7-Zip\7z.exe" a -tzip  xll.zip ".\%XLL_DIR%\*"
REM xcopy xll.zip "%VisualStudioDir%\%XLL_DIR%" /y