REM We need to launch the IDE with all the environment variables for building C++ apps.
REM https://stackoverflow.com/a/78884559/231860
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
cd "C:\Program Files\JetBrains\IntelliJ IDEA 2023.2.5\bin"
idea64.exe
