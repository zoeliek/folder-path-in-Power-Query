# Using parameter as folder path in Power Query
Replace the hardcoded folder path when connecting to SharePoint Folder from Power Query with parameter stored in Excel table.

## Problem 
Power Query relied on hardcoded folder path navigation when connecting to SharePoint to get data to Excel:  
When the folder name is changed or renamed, the pre-build query will fail and require a manual update in Advanced Editor in Power Query Editor.

## Objective
Create a table in one of the sheets in Excel to allow folder path to be updated directly, without requiring users to edit the code in Advanced Editor.

## Solution
**Step 1:**
Create a table in one of the sheets in Excel. The table will acts as user-maintained parameter.
| Parameter | File Path |
| -------- | -------- |
| TargetFolder | /Shared Documents/General/Data Collection Files/ |

Name the table to "Folder Path".  
Parameter is the name to capture the target folder path.  
File Path is the path of the file, normally in SharePoint, files are normally stored in "Shared Documents".

**Step 2:**  
Import the table into Power Query and create a parameter query. Update the query in Advanced Editor to
```powerquery
let
    Source = Excel.CurrentWorkbook(){[Name="FolderPath"]}[Content],
    TargetFolder = Source{0}[Value]
in
    TargetFolder
```
Then, save it.  

**Step 3:**
As usual step, in Excel, get data from "SharePoint Folder". Open the Power Query Editor, then update the code in Advanced Editor to
```powerquery
let
  TargetFolder = FolderPath,
  Source = SharePoint.Files("https://companyname.sharepoint.com/sites/itteam", [ApiVersion = 15]),
  FilteredFiles = Table.SelectRows(Source, each Text.Contains([Folder Path], TargetFolder))
in
  FilteredFiles
```
Then, locate the file and transform data (if needed). Save it.  
The folder path can be changed or updated in "Folder Path" table each time without opening the Power Query Editor.


