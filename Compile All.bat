@ECHO OFF

:: Calculate path to the Java lib. Make sure that the JAVA_HOME environment variable is set.
set JDK="%JAVA_HOME%\bin\javac.exe"

:: Paths to the Aspose.Words and other required libraries.
set LibPaths="..\..\..\lib\Aspose.Words.jdk15.jar;..\..\..\lib\demos-only-libs\testng-6.0.1.jar;..\..\..\lib\demos-only-libs\aspose-network.jar;..\..\..\lib\demos-only-libs\mail.jar;..\..\..\lib\demos-only-libs\activation.jar;"

:: Variables used by this script
set LogName=BuildLog.txt
set LogPath=..\..\%LogName%
set OutDir="out"

:: Counts
set /a SuccessCount=0
set /a ErrorCount=0

:: Calculate the current path and library path.
set BaseDir=%CD%
cd ..
set LibPath=%CD%\lib\
cd %BaseDir%

:: Abort if the JAVA_HOME variable is undefined.
IF "%JAVA_HOME%"=="" (
     echo Error: JAVA_HOME variable is not set. This should point to the directory where the JDK is installed to. Aborting.
     GOTO :end
)

:: Abort if the Java compiler cannot be found.
IF NOT EXIST %JDK% (
     echo Could not locate compiler at %JDK%. Aborting.
     GOTO :end
)

:: Abort if the lib folder cannot be found.
IF NOT EXIST "%LibPath%" ( 
    echo Error: Could not locate the lib folder at %LibPath%
    echo.
    echo Make sure you have copied the lib folder from the main Aspose.Words.Java.zip archive to this path.
    echo.
    GOTO :end
)

:: Check that each required library exists
call:checkLibraryExists "%LibPath%Aspose.Words.jdk15.jar"
call:checkLibraryExists "%LibPath%demos-only-libs\testng-6.0.1.jar"
call:checkLibraryExists "%LibPath%demos-only-libs\Aspose-Network.jar"
call:checkLibraryExists "%LibPath%demos-only-libs\mail.jar"
call:checkLibraryExists "%LibPath%demos-only-libs\activation.jar"

echo Compiling All Java Samples.
echo.

IF EXIST %LogName% del %LogName%

:: This will retrieve all sub folders at the current path and pass each one to compileSample function
:: To compile specific samples call compileSample FolderName. e.g call:compileSample AddWatermark.

for /F "delims=" %%j in ('dir /B /AD') do call:compileSample %%j

GOTO :success



:compileSample 
set sampleName=%~1

:: Skip potential folder if it does not contain a Java sub folder.
IF NOT EXIST %sampleName%\Java goto:eof

echo Compiling: %sampleName%
cd %sampleName%\Java

:: The out directory must already exist for javac
IF NOT EXIST %OutDir% md %OutDir%

echo ============================================================== >> %LogPath%
echo %sampleName% >>%LogPath%
echo ============================================================== >> %LogPath%
echo. >> %LogPath%

:: Redirect any errors to disk.
%JDK% -classpath %LibPaths% -d .\out -nowarn *.java >> %LogPath% 2>&1

if %ERRORLEVEL% == 0 (
echo            - Success.
echo. >> %LogPath%
echo Compiled Successfully >> %LogPath%
set /a SuccessCount+=1
) ELSE (
echo            - Build was unsuccessful, check the build log.
set /a ErrorCount+=1
)

echo.
echo. >> %LogPath%


:: Move back to the main samples directory.
cd ..\..
goto:eof


:checkLibraryExists
IF NOT EXIST "%~1" ( 
    echo Error: Could not locate the required library: 
    echo %~1
    echo.
    echo Make sure you have copied the lib folder from the main Aspose.Words.Java.zip archive to this path.
    echo.
    GOTO :end
)
goto:eof

:success
echo Build All Complete!
echo       - %SuccessCount% samples compiled successfully.
echo       - %ErrorCount% samples could not be compiled.

:end
pause
exit