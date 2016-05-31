set PATH=%PATH%;c:\Qt\5.5\msvc2013_64\bin
cls
windeployqt.exe --release --compiler-runtime ..\build-qxl-5_5_desktop_64bit_vs2013-Release\release\qxl.exe
pause
