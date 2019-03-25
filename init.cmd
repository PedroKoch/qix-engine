@ECHO OFF

ECHO Installing the necessary nodejs modules.
ECHO.

mkdir "./Backup"
mkdir "./Certificates"

SET dependencies=enigma.js ws xlsx path fs csv-writer

FOR %%i IN (%dependencies%) DO (
    ::ECHO "%%i"
    npm install %%i
    )

SET /p x="Press ENTER to finish"
ECHO.