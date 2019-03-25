@ECHO OFF

ECHO Backup master items from a qlik sense app.
ECHO.

Set /p app="Enter the AppID (Server) or App file path (Desktop): "
ECHO.

node src/auto.js Backup %app%

SET /p x="Press ENTER to finish"
ECHO.