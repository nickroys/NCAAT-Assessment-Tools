# NCAAT-Assessment-Tools
Collection of scripts to automate NCAAT Assessment Tasks

## Stig Details Creator:
Once you tun the checklists through Vulnerator, it will prodice an excel spreadsheet with all of the checklists in one sheet. In order to separate them into individual sheets, you can run the file through this script and it will find all of the individual file names (Column S) and proceed to create individual sheets named after each file with only the content from that checklist.

## Reauirements
1. Excel needs to be installed
2. The `STIG Details` sheet needs to be the name of the sheet. This is what Vulnerator defaults to so do not change it

## Options
You can edit the script to filter by a different column by changing the `$fileNameColIndex` variable. Currently it is set to 19 to use the file name, but in the event that you wanted to filter sheets by a different column you could do so by changing this variable
