Version 0.1 - It works, but still is pretty ugly in some senses
The tool was written mostly for the german community and by a german, so comments are german and it uses semicolons as delimiters in the created csv-files

I am using Read-HtmlTable by iRon 7: https://github.com/iRon7/Read-HtmlTable/tree/main in version 2.01 (the current version as of writing this)
Another part of the script was copied from this page: https://administrator.de/tutorial/powershell-einfuehrung-in-die-webbrowser-automation-mit-selenium-webdriver-1197173647.html - This will likely be changed significantly when overhauling the code

Known Issues:
- The script only functions reliably when the opened browser stays open, active and maximized during the parsing process. So please be patient for the minute or so this takes since I don't know if I can find a way around that
- Currently the script need to open a dummy browser (which is immediately closed again) to load the assemblies. This will need to be changed by moving functions around

Planned future improvements:
- Include some kind of interactive menu to let you select one of your characters and get a list of sessions / scenarios that character should have chronicles for
- Include functionality to list all scenarios you have played, with a differentiation between played and GMed
  - maybe even include "which ones am I missing"
- Introduce a way to load previously parsed data so that aforementioned functions can be run against the locally stored data with no need to parse again
- (Maybe) Introduce a way to send the data directly into a Google Sheet for further data anaylsis, etc. by the user
