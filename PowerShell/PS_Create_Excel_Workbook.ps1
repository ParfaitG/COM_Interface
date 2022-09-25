
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

    # CREATE PIVOT TABLE
    $pvtWks = $xlWbk.Worksheets.Add([System.Reflection.Missing]::Value, $xlWks)
    $pvtWks.Name = "PIVOT"

    $pvtCache = $xlWbk.PivotCaches().Create(1, $xlWks.UsedRange)
    $pvtTable = $pvtCache.CreatePivotTable($pvtWks.Cells(4, 2), "MetalsPivot")

    $pvtFld = $pvtTable.PivotFields("metal")
    $pvtFld.Orientation = 1
    $pvtFld.Position = 1

    $avgFld = $pvtTable.PivotFields("avg_price")
    $pvtTable.AddDataField($avgFld, "Min Price", -4139) | Out-Null
    $pvtTable.AddDataField($avgFld, "Avg Price", -4106) | Out-Null
    $pvtTable.AddDataField($avgFld, "Max Price", -4136) | Out-Null
    $pvtTable.AddDataField($avgFld, "Std Price", -4155) | Out-Null

    for($i = 3; $i -le 7; $i++){
        $rng = $pvtWks.Columns($i)
        $rng.NumberFormat = "$#,##0.00"
        $rng.HorizontalAlignment = -4152
    }

    # ADJUST DEFAULT FONT
    $xlCells = $pvtWks.Cells
    $xlCells.Font.Name = "Arial"
    $xlCells.Font.Size = "10"
    $xlCells.Font.Color = 0

    # SAVE WORKBOOK
    $xlWbk.SaveAs($xlFile, 51)

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
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($rng) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($avgFld) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pvtFld) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pvtTable) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pvtCache) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pvtWks) | Out-Null

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlCells) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlWks) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlWbk) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlApp) | Out-Null
}
