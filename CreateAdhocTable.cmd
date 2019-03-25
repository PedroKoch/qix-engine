@ECHO OFF

ECHO Create a master table with all master measures and dimensions in a qlik sense app.
ECHO.

Set /p app="Enter the AppID (Server) or App file path (Desktop): "
ECHO.

node src/auto.js CreateAdhoc %app%

SET /p x="Press ENTER to finish"
ECHO.