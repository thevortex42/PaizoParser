##### Version 0.1 - It works, but still is pretty ugly in some senses #####
DISCLAIMER: The tool was written mostly for the german community and by a german, so comments are german and it uses semicolons as delimiters in the created csv-files

This is a simple PowerShell tool that uses Selenium and the Edge Browser to open the Paizo.com website and get your Organized Play Characters as well as session history from there. The results are then displayed on screen, using Out-Gridview, as well as saved into CSV files, which you can then import into your favorite spreadsheet tool to use.
While designing the tool, I tried to make sure that only components that are present on any modern Windows computer are being used, so you don't need any advanced knowledge or install any software on your computer to use it.
No data is sent by this skript to any websites other than paizo.com and nuget.org, which hosts the drivers needed for Selenium to work.
Your user data is not being sent to any site other than paizo.com, and no data is collected by the tool in any way that I am aware of!

I am using Read-HtmlTable by iRon 7: https://github.com/iRon7/Read-HtmlTable/tree/main in version 2.01 (the current version as of writing this)
Another part of the script was copied from this page: https://administrator.de/tutorial/powershell-einfuehrung-in-die-webbrowser-automation-mit-selenium-webdriver-1197173647.html - This will likely be changed significantly when overhauling the code

Known Issues:
- The script only functions reliably when the opened browser stays open, active and maximized during the parsing process. So please be patient for the minute or so this takes since I don't know if I can find a way around that
- Currently the script need to open a dummy browser (which is immediately closed again) to load the assemblies. This will need to be changed by moving functions around

Planned future improvements:
- Change the csv-Delimiter to use local culture instead of defaulting to semicolon
- Change all comments / variable names / etc. to english
- Include some kind of interactive menu to let you select one of your characters and get a list of sessions / scenarios that character should have chronicles for
- Include functionality to list all scenarios you have played, with a differentiation between played and GMed
  - maybe even include "which ones am I missing"
- Introduce a way to load previously parsed data so that aforementioned functions can be run against the locally stored data with no need to parse again
- (Maybe) Introduce a way to send the data directly into a Google Sheet for further data anaylsis, etc. by the user
