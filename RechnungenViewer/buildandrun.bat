CALL "%VS140COMNTOOLS%\..\..\VC\vcvarsall.bat" x86_amd64

REM cd /d C:\Users\A.Roennburg\Documents\Visual Studio 2015\Projects\RechnungenViewer

msbuild /p:Configuration=debug
msbuild /p:Configuration=release
Debug\RechnungenViewer.exe