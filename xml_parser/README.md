# SQL XML Data Export

This PowerShell script gathers document records from SQL Server, then processes the files identified by those records.

## What the script does

1. Runs the configured SQL query to retrieve document records, including each file's path and filename.
2. Loops through the returned records and checks each file's type.
3. Extracts data from XML files and adds it to the CSV output.
4. Copies non-XML files to the configured additional-files output directory.

## Run the script

Configure the SQL connection, query, and output paths in `sql_to_xml_to_csv.ps1`, then run it from PowerShell:

```powershell
.\sql_to_xml_to_csv.ps1
```

The script writes the XML data to `xml_contracts.csv` in the configured output directory and copies other files to the configured additional-files directory.
