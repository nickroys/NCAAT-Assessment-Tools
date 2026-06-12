# NCAAT-Assessment-Tools
Collection of scripts to automate NCAAT Assessment Tasks

# Stig Details Creator
Once you run the checklists through Vulnerator, it will produce an excel spreadsheet with all of the checklists in one sheet. In order to separate them into individual sheets, you can run the file through this script and it will find all of the individual checklist names (Column S) and proceed to create individual sheets named after each one with only the content from that checklist.

## Requirements
1. Excel needs to be installed
2. Powershell execution policy should be set to `Bypass`
3. The `STIG Details` sheet needs to be named 'STIG Details'. This is what Vulnerator defaults to so do not change it

## Options
You can edit the script to filter by a different column by changing the `$fileNameColIndex` variable. Currently it is set to 19 to use the checklist name, but in the event that you wanted to filter sheets by a different column you could do so by changing this variable
