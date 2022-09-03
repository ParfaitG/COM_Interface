
##############################
### PATHS
##############################
$project_home = $Env:PROJECT_HOME
$csvData = "$($project_home)\PowerShell\Data\Precious_Metals_Prices.csv"
$xlFile = "$($project_home)\PowerShell\Outputs\PS_MSO_Spreadsheet.xlsx"


##############################
### CREATE EXCEL WORKBOOK
##############################
try {   
    # INITIALIZE COM OBJECT
    $xlApp = New-Object -ComObject Excel.Application
    $xlApp.DisplayAlerts = $false

    # CREATE WORKBOOK
    $xlWbk = $xlApp.Workbooks.Add()

    $xlWks = $xlWbk.Worksheets(1)
    $xlWks.Name = "METALS"

    # IMPORT CSV DATA
    $xlQt = $xlWks.QueryTables.Add("TEXT;$($csvData)", $xlWks.Range("A1"))

    $xlQt.TextFileParseType = 1
    $xlQt.TextFileCommaDelimiter = $true
    $xlQt.Refresh($false) | Out-Null
    $xlQt.Delete()

    # ADJUST DEFAULT FONT
    $xlCells = $xlWks.Cells
    $xlCells.Font.Name = "Arial"
    $xlCells.Font.Size = "10"
    $xlCells.Font.Color = 0
    
    # SAVE WORKBOOK
    $xlWbk.SaveAs($xlFile)

    # SHOW BACKGROUND APP
    $xlApp.Visible = $true
}
catch {
    Write-Host $_
    
    # CLOSE OBJECTS
    if ($xlQt) {
        $xlQt.Delete()
    }
    if ($xlWbk) {
        $xlWbk.Close($false)
    }
    if ($xlApp) {
        $xlApp.Quit()
    }
}
finally {
    # RELEASE RESOURCES
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlCells) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlWks) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlWbk) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlApp) | Out-Null
}


