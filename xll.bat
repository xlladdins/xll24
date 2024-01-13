REM Create project zip file for template
msbuild -t:Clean
set XLL_DIR=Templates\ProjectTemplates\Visual C++
del /s /q "%XLL_DIR%\xll\"
del /s /q "%XLL_DIR%\x64\"
del xll.zip
xcopy *.h "%XLL_DIR%\xll\" /y
xcopy *.c "%XLL_DIR%\xll\" /y
xcopy *.cpp "%XLL_DIR%\xll\" /y 
xcopy *.lib "%XLL_DIR%\xll\"  /y
xcopy xll.sln "%XLL_DIR%\xll\" /y
xcopy xll.vc* "%XLL_DIR%\xll\" /y
REM tar -C "%XLL_DIR%" -z -c -f xll.zip .
"C:\Program Files\7-Zip\7z.exe" a -tzip  xll.zip ".\%XLL_DIR%\*"
xcopy xll.zip "%VisualStudioDir%\%XLL_DIR%" /y