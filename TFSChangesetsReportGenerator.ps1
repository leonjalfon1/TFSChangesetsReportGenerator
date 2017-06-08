param
(
    [Parameter(Mandatory=$true)]
    $StartDate,
    [Parameter(Mandatory=$true)]
    $EndDate,
    [Parameter(Mandatory=$true)]
    $Path,
    [Parameter(Mandatory=$true)]
    $CollectionUrl,
    [Parameter(Mandatory=$true)]
    $TeamProject,
    [Parameter(Mandatory=$true)]
    $Credentials,
    [Parameter(Mandatory=$false)]
    $File = [Environment]::GetFolderPath("Desktop") + "\ChangesetsReport_" + (Get-Date).ToString('ddMMyyhhmmss') + ".html",
    [Parameter(Mandatory=$false)]
    $Detailed = $false
)

########################################################################################################################
# METHODS
########################################################################################################################


# Cast dates from "dd/mm/yyyy" to "mm/dd/yyyy"
function Cast-Date
{
    param
    (
        [Parameter(Mandatory=$true)]
        $Date
    )

    try
    {
        $day = $Date.Split("/")[0]
        $month = $Date.Split("/")[1]
        $year = $Date.Split("/")[2]
        $NewDate = $month + "/" + $day + "/" + $year
        return $NewDate
    }
    catch
    {
        Write-Output "Invalid Date, Exception: $_.Exception.Message"
        Exit 101
    }
}


# Get BuidDefinitionId using BuildDefinitionName
function Get-ChangesetsByDates
{
    param
    (
        [Parameter(Mandatory=$true)]
        $StartDate,
        [Parameter(Mandatory=$true)]
        $EndDate,
        [Parameter(Mandatory=$true)]
        $Path,
        [Parameter(Mandatory=$true)]
        $File,
        [Parameter(Mandatory=$true)]
        $CollectionUri,
        [Parameter(Mandatory=$true)]
        $TeamProject,
        [Parameter(Mandatory=$true)]
        $Credentials
    )

    $Page = 0
    
    try
    {
        $TempFile = $File.Substring(0,$File.Length-5)+"_temp.txt"
        if(Test-Path $TempFile) {Remove-Item $TempFile}

        $count = 0

        Do
        {
            $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}" -f $Credentials)))
            $Response = Invoke-WebRequest -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -ContentType application/json -Uri "$CollectionUri/$TeamProject/_apis/tfvc/changesets?searchCriteria.fromDate=$StartDate&searchCriteria.toDate=$EndDate&searchCriteria.itemPath=$Path&`$top=100&`$skip=$Page&api-version=1.0" -Method GET -UseBasicParsing
            $ResponseJson = $Response | ConvertFrom-Json
            $TotalChangesets = $ResponseJson.count
            $ChangesetsResults = $ResponseJson.value

            $count = $count + $TotalChangesets
            Write-Host "$count changesets found"

            foreach($Changeset in $ChangesetsResults)
            {
                Add-Content $TempFile $Changeset.changesetId | out-null
            }

            $Page = $Page + 1
        }
        While($TotalChangesets -eq 100)

        return $TempFile

    }
    catch
    {
        Write-Host "Failed to get changesets, Exception: $_.Exception.Message"
        Exit 102
    }
}


# Get ChangesetInfo from ChangesetId
function Get-ChangesetInfo
{
    param
    (
        [Parameter(Mandatory=$true)]
        $ChangesetId,
        [Parameter(Mandatory=$true)]
        $CollectionUri,
        [Parameter(Mandatory=$true)]
        $TeamProject,
        [Parameter(Mandatory=$true)]
        $Credentials
    )

    try
    {
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}" -f $Credentials)))
        $RequestUrl = "$CollectionUri/$TeamProject/_apis/tfvc/changesets/" + $ChangesetId + "?api-version=1.0"
        $Response = Invoke-WebRequest -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -ContentType application/json -Uri $RequestUrl -Method GET -UseBasicParsing
        $ChangesetInfo = $Response | ConvertFrom-Json
        return $ChangesetInfo
    }
    catch
    {
        Write-Output "Failed to get the ChangesetInfo, Exception: $_.Exception.Message"
        Exit 103
    }
}


# Create Report per half hour
function Create-DetailedReport
{
    param
    (
        [Parameter(Mandatory=$true)]
        $StartDate,
        [Parameter(Mandatory=$true)]
        $EndDate,
        [Parameter(Mandatory=$true)]
        $Path,
        [Parameter(Mandatory=$true)]
        $File,
        [Parameter(Mandatory=$true)]
        $Data,
        [Parameter(Mandatory=$true)]
        $TotalChangesets
    )

    try
    {
        if(Test-Path $File) {Remove-Item $File}
 
        Add-Content $File '<html>'
        Add-Content $File '  <head>'
        Add-Content $File '    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>'
        Add-Content $File '    <script type="text/javascript">'
        Add-Content $File "      google.charts.load('current', {'packages':['bar']});"
        Add-Content $File '      google.charts.setOnLoadCallback(drawStuff);'
        Add-Content $File ''
        Add-Content $File '      function drawStuff() {'
        Add-Content $File '        var data = new google.visualization.arrayToDataTable(['
        Add-Content $File "         ['Move', 'Percentage'],"
        $Line = '          ["00:00", ' + $Data[0] + '],'
        Add-Content $File $Line
        $Line = '          ["00:30", ' + $Data[1] + '],'
        Add-Content $File $Line
        $Line = '          ["01:00", ' + $Data[2] + '],'
        Add-Content $File $Line
        $Line = '          ["01:30", ' + $Data[3] + '],'
        Add-Content $File $Line
        $Line = '          ["02:00", ' + $Data[4] + '],'
        Add-Content $File $Line
        $Line = '          ["02:30", ' + $Data[5] + '],'
        Add-Content $File $Line
        $Line = '          ["03:00", ' + $Data[6] + '],'
        Add-Content $File $Line
        $Line = '          ["03:30", ' + $Data[7] + '],'
        Add-Content $File $Line
        $Line = '          ["04:00", ' + $Data[8] + '],'
        Add-Content $File $Line
        $Line = '          ["04:30", ' + $Data[9] + '],'
        Add-Content $File $Line
        $Line = '          ["05:00", ' + $Data[10] + '],'
        Add-Content $File $Line
        $Line = '          ["05:30", ' + $Data[11] + '],'
        Add-Content $File $Line
        $Line = '          ["06:00", ' + $Data[12] + '],'
        Add-Content $File $Line
        $Line = '          ["06:30", ' + $Data[13] + '],'
        Add-Content $File $Line
        $Line = '          ["07:00", ' + $Data[14] + '],'
        Add-Content $File $Line
        $Line = '          ["07:30", ' + $Data[15] + '],'
        Add-Content $File $Line
        $Line = '          ["08:00", ' + $Data[16] + '],'
        Add-Content $File $Line
        $Line = '          ["08:30", ' + $Data[17] + '],'
        Add-Content $File $Line
        $Line = '          ["09:00", ' + $Data[18] + '],'
        Add-Content $File $Line
        $Line = '          ["09:30", ' + $Data[19] + '],'
        Add-Content $File $Line
        $Line = '          ["10:00", ' + $Data[20] + '],'
        Add-Content $File $Line
        $Line = '          ["10:30", ' + $Data[21] + '],'
        Add-Content $File $Line
        $Line = '          ["11:00", ' + $Data[22] + '],'
        Add-Content $File $Line
        $Line = '          ["11:30", ' + $Data[23] + '],'
        Add-Content $File $Line
        $Line = '          ["12:00", ' + $Data[24] + '],'
        Add-Content $File $Line
        $Line = '          ["12:30", ' + $Data[25] + '],'
        Add-Content $File $Line
        $Line = '          ["13:00", ' + $Data[26] + '],'
        Add-Content $File $Line
        $Line = '          ["13:30", ' + $Data[27] + '],'
        Add-Content $File $Line
        $Line = '          ["14:00", ' + $Data[28] + '],'
        Add-Content $File $Line
        $Line = '          ["14:30", ' + $Data[29] + '],'
        Add-Content $File $Line
        $Line = '          ["15:00", ' + $Data[30] + '],'
        Add-Content $File $Line
        $Line = '          ["15:30", ' + $Data[31] + '],'
        Add-Content $File $Line
        $Line = '          ["16:00", ' + $Data[32] + '],'
        Add-Content $File $Line
        $Line = '          ["16:30", ' + $Data[33] + '],'
        Add-Content $File $Line
        $Line = '          ["17:00", ' + $Data[34] + '],'
        Add-Content $File $Line
        $Line = '          ["17:30", ' + $Data[35] + '],'
        Add-Content $File $Line
        $Line = '          ["18:00", ' + $Data[36] + '],'
        Add-Content $File $Line
        $Line = '          ["18:30", ' + $Data[37] + '],'
        Add-Content $File $Line
        $Line = '          ["19:00", ' + $Data[38] + '],'
        Add-Content $File $Line
        $Line = '          ["19:30", ' + $Data[39] + '],'
        Add-Content $File $Line
        $Line = '          ["20:00", ' + $Data[40] + '],'
        Add-Content $File $Line
        $Line = '          ["20:30", ' + $Data[41] + '],'
        Add-Content $File $Line
        $Line = '          ["21:00", ' + $Data[42] + '],'
        Add-Content $File $Line
        $Line = '          ["21:30", ' + $Data[43] + '],'
        Add-Content $File $Line
        $Line = '          ["22:00", ' + $Data[44] + '],'
        Add-Content $File $Line
        $Line = '          ["22:30", ' + $Data[45] + '],'
        Add-Content $File $Line
        $Line = '          ["23:00", ' + $Data[46] + '],'
        Add-Content $File $Line
        $Line = '          ["23:30", ' + $Data[47] + ']'
        Add-Content $File $Line
        Add-Content $File '        ]);'
        Add-Content $File ''
        Add-Content $File '        var options = {'
        Add-Content $File '          width: 1800,'
        Add-Content $File "          legend: { position: 'none' },"
        Add-Content $File '          chart: {'
        Add-Content $File "            title: 'TFS Changesets per Hour',"
        $Line = '            subtitle: ' + "'" + $Path + " \n " + $StartDate.Replace("-","/") + ' - ' + $EndDate.Replace("-","/") + "'" + ' },'
        Add-Content $File $Line
        Add-Content $File '          axes: {'
        Add-Content $File '            x: {'
        $Line = "              0: { side: 'bottom', label: 'Total Changesets: " + $TotalChangesets + "'} // Top x-axis."
        Add-Content $File $Line
        Add-Content $File '            }'
        Add-Content $File '          },'
        Add-Content $File '          bar: { groupWidth: "90%" }'
        Add-Content $File '        };'
        Add-Content $File ''
        Add-Content $File "        var chart = new google.charts.Bar(document.getElementById('top_x_div'));"
        Add-Content $File '        // Convert the Classic options to Material options.'
        Add-Content $File '        chart.draw(data, google.charts.Bar.convertOptions(options));'
        Add-Content $File '      };'
        Add-Content $File '    </script>'
        Add-Content $File '  </head>'
        Add-Content $File '  <body>'
        Add-Content $File '    <div id="top_x_div" style="width: 800px; height: 600px;"></div>'
        Add-Content $File '  </body>'
        Add-Content $File '</html>'
    }
    catch
    {
        Write-Output "Failed creating the report, Exception: $_.Exception.Message"
        Exit 104
    }
}


# Create Report per hour
function Create-Report
{
    param
    (
        [Parameter(Mandatory=$true)]
        $StartDate,
        [Parameter(Mandatory=$true)]
        $EndDate,
        [Parameter(Mandatory=$true)]
        $Path,
        [Parameter(Mandatory=$true)]
        $File,
        [Parameter(Mandatory=$true)]
        $Data,
        [Parameter(Mandatory=$true)]
        $TotalChangesets
    )

    try
    {
        if(Test-Path $File) {Remove-Item $File}
 
        Add-Content $File '<html>'
        Add-Content $File '  <head>'
        Add-Content $File '    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>'
        Add-Content $File '    <script type="text/javascript">'
        Add-Content $File "      google.charts.load('current', {'packages':['bar']});"
        Add-Content $File '      google.charts.setOnLoadCallback(drawStuff);'
        Add-Content $File ''
        Add-Content $File '      function drawStuff() {'
        Add-Content $File '        var data = new google.visualization.arrayToDataTable(['
        Add-Content $File "         ['Move', 'Percentage'],"
        
        [int]$Total = $Data[0] + $Data[1]
        $Line = '          ["00:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[2] + $Data[3]
        $Line = '          ["01:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[4] + $Data[5]
        $Line = '          ["02:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[6] + $Data[7]
        $Line = '          ["03:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[8] + $Data[9]
        $Line = '          ["04:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[10] + $Data[11]
        $Line = '          ["05:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[12] + $Data[13]
        $Line = '          ["06:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[14] + $Data[15]
        $Line = '          ["07:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[16] + $Data[17]
        $Line = '          ["08:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[18] + $Data[19]
        $Line = '          ["09:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[20] + $Data[21]
        $Line = '          ["10:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[22] + $Data[23]
        $Line = '          ["11:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[24] + $Data[25]
        $Line = '          ["12:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[26] + $Data[27]
        $Line = '          ["13:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[28] + $Data[29]
        $Line = '          ["14:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[30] + $Data[31]
        $Line = '          ["15:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[32] + $Data[33]
        $Line = '          ["16:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[34] + $Data[35]
        $Line = '          ["17:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[36] + $Data[37]
        $Line = '          ["18:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[38] + $Data[39]
        $Line = '          ["19:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[40] + $Data[41]
        $Line = '          ["20:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[42] + $Data[43]
        $Line = '          ["21:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[44] + $Data[45]
        $Line = '          ["22:00", ' + $Total + '],'
        Add-Content $File $Line

        [int]$Total = $Data[46] + $Data[47]
        $Line = '          ["23:00", ' + $Total + '],'
        Add-Content $File $Line

        Add-Content $File '        ]);'
        Add-Content $File ''
        Add-Content $File '        var options = {'
        Add-Content $File '          width: 1200,'
        Add-Content $File "          legend: { position: 'none' },"
        Add-Content $File '          chart: {'
        Add-Content $File "            title: 'TFS Changesets per Hour',"
        $Line = '            subtitle: ' + "'" + $Path + " \n " + $StartDate.Replace("-","/") + ' - ' + $EndDate.Replace("-","/") + "'" + ' },'
        Add-Content $File $Line
        Add-Content $File '          axes: {'
        Add-Content $File '            x: {'
        $Line = "              0: { side: 'bottom', label: 'Total Changesets: " + $TotalChangesets + "'} // Top x-axis."
        Add-Content $File $Line
        Add-Content $File '            }'
        Add-Content $File '          },'
        Add-Content $File '          bar: { groupWidth: "90%" }'
        Add-Content $File '        };'
        Add-Content $File ''
        Add-Content $File "        var chart = new google.charts.Bar(document.getElementById('top_x_div'));"
        Add-Content $File '        // Convert the Classic options to Material options.'
        Add-Content $File '        chart.draw(data, google.charts.Bar.convertOptions(options));'
        Add-Content $File '      };'
        Add-Content $File '    </script>'
        Add-Content $File '  </head>'
        Add-Content $File '  <body>'
        Add-Content $File '    <div id="top_x_div" style="width: 800px; height: 600px;"></div>'
        Add-Content $File '  </body>'
        Add-Content $File '</html>'
}
    catch
    {
        Write-Output "Failed creating the report, Exception: $_.Exception.Message"
        Exit 105
    }
}


########################################################################################################################
# GENERATE REPORT
########################################################################################################################

# Cast dates to TFS format
$OriginalStartDate = $StartDate
$OriginalEndDate = $EndDate
$StartDate = Cast-Date -Date $StartDate
$EndDate = Cast-Date -Date $EndDate

# Define variables
$count = 1
$FinalTable = 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0

# Get changesets
$ChangesetsFile = Get-ChangesetsByDates -StartDate $StartDate -EndDate $EndDate -Path $Path -File $File -CollectionUri $CollectionUrl -TeamProject $TeamProject -Credentials $Credentials
$Changesets = Get-Content $ChangesetsFile

# Build array with statistics
foreach($ChangesetId in $Changesets)
{
   # get checkin time
   $ChangesetInfo = Get-ChangesetInfo -ChangesetId $ChangesetId -CollectionUri $CollectionUrl -TeamProject $TeamProject -Credentials $Credentials

   # write status in console
   $StatusMessage = "Retrieving data from changeset #$ChangesetId ($count of " + $Changesets.Count + ")"
   Write-Host $StatusMessage
   $count++

   # get creation type and sort
   $ChangesetCreationTime = $ChangesetInfo.createdDate.Split("T")[1].Split(".")[0].Substring(0,5)
   $Hour = [convert]::ToInt32($ChangesetCreationTime.Split(":")[0], 10)
   $Minutes = [convert]::ToInt32($ChangesetCreationTime.Split(":")[1], 10)

   if($Minutes -gt 29)
   {
      $Hour = $Hour + 0.5
   }

   # add to the final "table"
   $FinalTable[$Hour*2] = $FinalTable[$Hour*2]+1
}


# Delete Temp file

if(Test-Path $ChangesetsFile) {Remove-Item $ChangesetsFile}


# Create report and open it

if($Detailed)
{
    Create-DetailedReport -StartDate $OriginalStartDate -EndDate $OriginalEndDate -Path $Path -File $File -Data $FinalTable -TotalChangesets $Changesets.Count
}
else
{
    Create-Report -StartDate $OriginalStartDate -EndDate $OriginalEndDate -Path $Path -File $File -Data $FinalTable -TotalChangesets $Changesets.Count
}

Write-Host "Report available in: {$File}"
Start-Process "chrome.exe" $File


##############################################################################################