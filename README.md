# NCAAT-Assessment-Tools
Collection of scripts to automate NCAAT Assessment Tasks

# Stig Details Creator
Once you run the checklists through Vulnerator, it will produce an excel spreadsheet with all of the checklists in one sheet. In order to separate them into individual sheets, you can run the file through this script and it will find all of the individual checklist names (Column S) and proceed to create individual sheets named after each one with only the content from that checklist.

## Requirements
1. Excel needs to be installed
3. The `STIG Details` sheet needs to be named 'STIG Details'. This is what Vulnerator defaults to so do not change it

## Running the tool
Powershell will likely stop this script from running if the execution policy is not set to `Bypass`

In order to run the script without permaently, open a Powershell terminal in the directory that the script is stored and use this command:
```powershell.exe -ExecutionPolicy Bypass -File .\STIG_Details_Creator_Interactive.ps1```

Then after it runs, double check that your execution policy has not channged:
```Get-ExecutionPolicy```

Leaving your Powershell execution policy set to `Bypass` is bad OPSEC and should be avoided at all costs

## Options
You can edit the script to filter by a different column by changing the `$fileNameColIndex` variable. Currently it is set to 19 to use the checklist name, but in the event that you wanted to filter sheets by a different column you could do so by changing this variable
