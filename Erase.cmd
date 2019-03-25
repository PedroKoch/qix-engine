@ECHO OFF

ECHO Erase master measures, master dimensions and variables from a qlik sense app.
ECHO.

Set /p app="Enter the AppID (Server) or App file path (Desktop): "
ECHO.

node src/auto.js EraseAll %app%

SET /p x="Press ENTER to finish"
ECHO.