# Gas_Scan_Bill

This script processes gas bills, collect data from them and update the utility spreadsheet.

Below is a detailed explanation of how the overall system work.

## How it works

1. Create a folder for the month you are interested in
2. Get the most recent version of the spreadsheet
3. Go to Center Point Website
4. Download all the bills and put them in the folder created
5. Go to the command line and enter the following command

-f : folder with all the bills

-e : excel sheet to update

-m : month abbreviation

```
$ python3 electricityScan.py -f 'folder_name' -e 'spreadsheet_name' -m 'month'
```

6. Example for April. Folder called april and month abbreviation is Apr.

```
$ python3 gasScan.py -f april -e "2024-25 Utility Billing.xlsx" -m 'Apr.'
```

7. A spreadsheet called 2024-25 Utility Billing Output.xlsx should be the output

8. Update the spreadsheet in Box (copy and paste or update the entire spreadsheet)
