REM Create project zip file for template
set XLL_DIR=Templates\ProjectTemplates\Visual C++
rmdir /s /q "%XLL_DIR%\.vs"
rmdir /s /q "%XLL_DIR%\x64"
del xll.zip
tar -c -z xll.zip -C "%XLL_DIR%"
REM "C:\Program Files\7-Zip\7z.exe" a -tzip  xll.zip ".\%XLL_DIR%\*"
REM Install xll template for Visual Studio
xcopy xll.zip "%VisualStudioDir%\%XLL_DIR%\" /y