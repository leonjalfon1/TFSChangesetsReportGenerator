# TFSChangesetsReportGenerator
Powershell script to generate a TFS changesets report

---

## Report

 - Retrieve the total of changesets between two dates in a specific path
 - Shows the number of checkins per hour (or half hour)

---
 
## Pre-Requisites

 - Powershell version 4.0 or later
 - Connection to internet (for access the libraries to generate the chart)
 - Access to the TFS Rest API
 
---
 
## Parameters

### Mandatory

#### StartDate
 - Date and time (optional) of the earliest changeset to return
 - Example: "6/7/2017"
#### EndDate
 - Date and time of the latest changesets to return
 - Example: "6/8/2017 11:59:59 PM"
#### Path
 - Path in source control (directory or file)
 - Example: "$/TeamProject/Branch/Directory"
#### CollectionUrl
 - TFS Collection Url
 - Example: "http://tfsserver:8080/tfs/DefaultCollection"
#### TeamProject
 - TFS Team Project Name
 - Example: "MyProject"
#### Credentials
 - Credentials to use to connect with the TFS REST Api
 - Example: "Domain\User:Password"

 
### Optional

#### File
 - Specify the file where the summary will be created (the default is "Desktop\ChangesetsByHourReport_{datetime}.html")
 - Example: "C:\Users\User\Desktop\TFSChangesetsReportGenerator\Report.html"
#### Detailed
 - Indicates if show the results per hour or per half hours
 - $false by default, show total changesets per hour (valid values: $true or $false)
 - Example: $true

---

## Usage

- Mandatory
```
.\TFSChangesetsReportGenerator.ps1 -StartDate "6/7/2017" -EndDate "6/8/2017 11:59:59 PM" -Path "$/TP/Dev" -CollectionUrl "http://tfsserver:8080/tfs/DefaultCollection" -TeamProject "MyProject" -Credentials "dom\leonj:12345"
```
- All
```
.\TFSChangesetsReportGenerator.ps1 -StartDate "6/7/2017" -EndDate "6/8/2017 11:59:59 PM" -Path "$/TP/Dev" -CollectionUrl "http://tfsserver:8080/tfs/DefaultCollection" -TeamProject "MyProject" -Credentials "dom\leonj:12345" -File "C:\Users\User\Desktop\TFSChangesetsReportGenerator\Report.html" -Detailed $true
```
 
---
 
## Changeset Report

 - Basic Report Example: [ChangesetsReportExample](ChangesetsReportExample.html)
 - Detailed Report Example: [ChangesetsDetailedReportExample](ChangesetsDetailedReportExample.html)
 
---

## Contributing

 - Please feel free to contribute, suggest ideas or open issues
