

##############################
### PATHS
##############################
$project_home = $Env:PROJECT_HOME
$aggData = "$($project_home)\PowerShell\Data\Precious_Metals_Prices_Summary.csv"
$boxplotImg = "$project_home\PowerShell\Data\Precious_Metals_BoxPlot.png"
$yearplotImg = "$project_home\PowerShell\Data\Precious_Metals_YearPlot.png"


##############################
### DATA
##############################

$agg_data = Import-Csv -Path $aggData
$agg_cols = $agg_data[0].psobject.properties.name


# COPY ATTACHMENT SOURCE FILES FOR EMAIL
function New-TemporaryDirectory {
    $parent = [System.IO.Path]::GetTempPath()
    [string] $name = [System.Guid]::NewGuid()
    New-Item -ItemType Directory -Path (Join-Path $parent $name)
}
$tmpDir = New-TemporaryDirectory

Copy-Item -Path "$project_home\PowerShell\Outputs\*" `
          -Destination $tmpDir `
          -Force `
          -Recurse

Remove-Item "$tmpDir\*laccdb"
$msoFiles = @(Get-ChildItem -Path $tmpDir)


##############################
### CREATE OUTLOOK EMAIL
##############################
try {
    # INITIALIZE COM OBJECT
    $olApp = New-Object -ComObject Outlook.Application

    # CREATE EMAIL
    $olMail = $olApp.CreateItem(0)

    # COLUMNS
    $olHdr = "<tr>"
    for($j = 0; $j -le 5; $j++) {
        $olHdr = "$olHdr<th>$($agg_cols[$j])</th>"
    }
    $olHdr = "$olHdr</tr>"

    # ROWS
    $olTbl = ""
    for($i=2; $i -lt 6; $i++){
        $olTbl = "$olTbl<tr>"
        for($j=0; $j -le 5; $j++){
            $olTbl = (
                "$olTbl<td>" +
                $agg_data[$i-2].PSObject.Properties[$agg_cols[$j]].value +
                "</td>"
            )
        }
        $olTbl = "$olTbl</tr>"
    }

    # ADD RECIPIENTS AND SUBJECT
    $insp = $olMail.GetInspector
    $signature = $olMail.HTMLBody
    $olMail.Recipients.Add("crug@meetup.com") | Out-Null
    $olMail.Subject = "Precious Metals Analysis"

    # ADD ATTACHMENTS
    $olAttachments = $olMail.Attachments
    ForEach($msoFile in $msoFiles) {
        $olAttachments.Add($msoFile.FullName, 1) | Out-Null
    }
    $olAttachments.Add($boxplotImg, 1, 0) | Out-Null
    $olAttachments.Add($yearplotImg, 1, 0) | Out-Null

    # ADD BODY
    $olMail.HTMLBody = (
      '<html>' +
      '<head>' +
      '<style type="text/css">' +
      '   .aggtable {font-family: Arial; font-size: 14px; border-collapse: collapse;}' +
      '   .aggtable th, .aggtable td {border: 1px solid #CCC; text-align: right; padding: 4px;}' +
      '   .aggtable tr:nth-child(even) {background-color: #f2f2f2;}' +
      '</style>' +
      '</head>' +
      '<body style="font-family: Arial; font-size: 14px;">' +
      '<p>Hello CRUG useRs!<br/><br/>' +
      '<p>Please find below our analysis of precious metals, powered by R!</p>' +
      '<br/><div style="text-align: center;"/>' +
      '<table class="aggtable">' +
      $olHdr +
      $olTbl +
      '</table><br/>' +
      '<img src="cid:Precious_Metals_BoxPlot.png" width="960"/><br/>' +
      '<img src="cid:Precious_Metals_YearPlot.png" width="960"/>' +
      '</div>' +
      $signature +
      '</p>' +
      '</body>' +
      '</html>'
    )

    # DISPLAY EMAIL
    $olMail.Display() | Out-Null
}
catch {
    Write-Host $_

    # CLOSE OBJECTS
    if ($olMail) {
        $olMail.Close($false)
    }
    if ($olApp) {
        $olApp.Quit()
    }
}
finally {
    # RELEASE RESOURCES
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($olMail) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($olApp) | Out-Null
}
