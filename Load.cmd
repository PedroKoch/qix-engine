@ECHO OFF

ECHO Load master items from an excel file to a qlik sense app.
ECHO.

Set /p app="Enter the AppID (Server) or App file path (Desktop): "
Set /p file="Enter the Excel file name: "
ECHO.

node src/auto.js Load %app% %file%

SET /p x="Press ENTER to finish"
ECHO.