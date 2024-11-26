# Archive Notice
This is now archived as MSF Tools is obsolete

# msftools-groot-import
A Google Apps script to import data from the MSF Toolbot into the G.R.O.O.T. with one click

## Adding this script to your G.R.O.O.T.

- Create a new sheet (not spreadsheet) in the G.R.O.O.T. called `*Toolbot`. Ensure that there is an "\*" at the beginning of `*Toolbot`. This makes it so that if you upgrade to a new G.R.O.O.T., it will be copied over.
- ![Add the \*Toolbot Sheet](/.images/add-sheet.png)
- ![Name the new sheet \*Toolbot](/.images/name-toolbot.png)
- In cell A1 in `*Toolbot`, enter your Toolbot Sheet ID
- ![Add your MSf ToolBot sheet ID](/.images/sheet-id.png)
- Add the script to the script collection
  - Navigate to Tools > Script Editor
  - ![Open the Script Editor](/.images/script-editor.png)
  - File > New > Script file. Name it whatever you like
  - ![Add a new Script File](/.images/new-script.png)
  - ![Name the script file](/.images/name-script.png)
  - Copy and paste the toolbot_import.gs contents into the new script file
  - ![Copy and paste the script](/.images/script-contents.png)
  - Add a menu item to Triggers.gs: `topMenu.addItem("Import from MSF Toolbot", "msftoolbot_import");` on the line before `topMenu.addToUi();`
  - ![Add a menu item](/.images/add-menu.png)
- Save and refresh your G.R.O.O.T. (This will close the Script Editor). Use the browser refresh button or F5. If you use ctrl+R, it doesn't refresh the entire page, just the contents inside.

- Access the newly added script under Zaratools > Import from MSF Toolbot (If you do not see the Zaratools custom menu, you may have to refresh the page again
- ![Use the script](/.images/use-script.png)

## Credits

Discord user Hazelwood#9033 for creating the MSF ToolBot. [MSF ToolBot Discord](https://discord.gg/mXRjJp3)

Discord user Zarathos#6866 for creating the G.R.O.O.T. [Zaratools Discord](https://discord.gg/vqhkZEp)
