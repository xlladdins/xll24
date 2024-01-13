REM Create project zip file for template
msbuild -t:Clean
set XLL_DIR="Templates\ProjectTemplates\Visual C++"
del /s "%XLL_DIR%\xll\"
del "%XLL_DIR%\xll.zip"
xcopy *.h "%XLL_DIR%\xll\" /y
xcopy *.c "%XLL_DIR%\xll\" /y
xcopy *.cpp "%XLL_DIR%\xll\" /y 
xcopy *.lib "%XLL_DIR%\xll\"  /y
xcopy xll.sln "%XLL_DIR%\xll\" /y
xcopy xll.vc* "%XLL_DIR%\xll\" /y
cd "%XLL_DIR%"
tar -a -z -c -f xll.zip .
xcopy xll.zip "%VisualStudioDir%\%XLL_DIR%" /y