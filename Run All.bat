@ECHO OFF
setLocal EnableDelayedExpansion

:: Paths to the Aspose.Words library and other required libraries.
set LibPaths="..\..\..\..\lib\Aspose.Words.jdk15.jar;..\..\..\..\lib\demos-only-libs\testng-6.0.1.jar;..\..\..\..\lib\demos-only-libs\aspose-network.jar;..\..\..\lib\demos-only-libs\mail.jar;..\..\..\lib\demos-only-libs\activation.jar;"

:: Counts
set /a SuccessCount=0
set /a ErrorCount=0
set /a SkippedCount=0

:: Run dummy execution of java.exe which should exist in the PATH variable to test if it exists.
java > NUL 2> NUL 
if NOT %ERRORLEVEL%==0 (
     echo Could not locate the java runtime. Ensure that the appropriate path exists to java.exe in the PATH variable.
     GOTO :end
)

echo Executing All Java Samples.
echo.

:: This will retrieve all sub folders at the current path and pass each one to the runSample function
:: To compile specific samples call runSample FolderName. e.g call:runSample AddWatermark.

for /F "delims=" %%j in ('dir /B /AD') do call:runSample %%j

GOTO :success

:runSample
set sampleName=%~1

:: Skip these particular samples as they cannot be properly run through the command line.
IF "%sampleName%" == "SaveHtmlAndEmail" goto :eof
IF "%sampleName%" == "DocumentPreviewAndPrint" goto :eof
IF "%sampleName%" == "MultiplePagesOnSheet" goto :eof

:: Skip potential folder if it does not contain a Java sub folder.
IF NOT EXIST %sampleName%\Java goto:eof

:: If the generated folder contains no class files then assume the project could not be compiled.
:: Skip this sample.
IF NOT EXIST %sampleName%\Java\out\%sampleName%\*.class (
  set /a SkippedCount+=1
  echo     - Could not find class file in %sampleName%. Skipping sample.
  goto:eof
)

cd %sampleName%\Java\out\

echo Running: %sampleName%

IF "%sampleName%"=="Examples" (

 :: Create a comma separated list of the class files in the directory.
 for /f "tokens=* delims= " %%a in ('dir/b/a-d Examples\*.class') do (
  set /a N+=1
  if !N! equ 1 ( set str=Examples.%%a
  ) else (
  set str=!str!,Examples.%%a
  )
 )
 :: Pass these classes to be run by the unittest manager.
 java -classpath %LibPaths% org.testng.TestNG -testclass !str!
) ELSE (
 :: Run this sample normally.
 java -ea -classpath %LibPaths% %sampleName%.Program
)

:: Check the result of the program and count if it executed successfully or an error occured.
IF %ERRORLEVEL% == 0 (
 set /a SuccessCount+=1
 echo.
 echo     - Completed Successfully.
) ELSE (
 set /a ErrorCount+=1
 echo.
 echo     - Error encountered.
)

cd ..\..\..

echo.
goto:eof

:success
echo Run All Complete!
echo     - %SuccessCount% samples run successfully.
echo     - %ErrorCount% samples encountered an exception.
echo     - %SkippedCount% samples which are not compiled and were skipped.

:end
pause