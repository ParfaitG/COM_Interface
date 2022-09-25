
##############################
### PATHS
##############################
$project_home = $Env:PROJECT_HOME
$csvData = "$($project_home)\PowerShell\Data\Precious_Metals_Prices.csv"
$dbFile = "$($project_home)\PowerShell\Outputs\PS_MSO_Database.accdb"


##############################
### CREATE ACCESS DATABASE
##############################
try {    
    # INITIALIZE COM OBJECT
    $accApp = New-Object -ComObject Access.Application
    
    
    if (!(Test-Path -path $dbFile)) {
        # CREATE DATABASE
        $accApp.NewCurrentDatabase($dbFile) | Out-Null
    }
    else{
        # OPEN DATABASE
        $accApp.OpenCurrentDatabase($dbFile) | Out-Null
    }

    $accDB = $accApp.CurrentDb()

    # CLEAN OUT TABLE
    foreach($tbl in $accDB.TableDefs) {
        if ($tbl.Name -eq "metals") {            
            $accApp.CurrentDb().Execute("DELETE FROM metals")
        }
    }
    
    # IMPORT CSV
    $accApp.DoCmd.TransferText(0, $null, "metals", $csvData, $true)

    # CREATE QUERY
    foreach($qry in $accDB.QueryDefs) {
        if ($qry.Name -eq "metals_agg") {            
            $accApp.CurrentDb().Execute("DROP TABLE metals_agg")
        }
    }
    
    $qry = $accDB.CreateQueryDef(
        "metals_agg",
        (
            "SELECT metal, " +
            "       Ccur(MIN(avg_price)) AS MinPrice, " +
            "       Ccur(AVG(avg_price)) AS AvgPrice, " +
            "       Ccur(MAX(avg_price)) AS MaxPrice, " +
            "       Ccur(STDEV(avg_price)) AS StdPrice " +
            "FROM metals " +
            "GROUP BY metal"
        )
    )
    
    #SHOW BACKGROUND APP
    $accApp.UserControl = $true
    $accApp.DoCmd.OpenTable("metals") | Out-Null
    $accApp.DoCmd.OpenQuery("metals_agg") | Out-Null
}
catch {
    Write-Host $_

    # CLOSE OBJECT
    if ($accApp) {
        $accApp.DoCmd.CloseDatabase()
        $accApp.Quit()
    }
}
finally {
    # RELEASE RESOURCES
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($tbl) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($qry) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($accDB) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($accApp) | Out-Null
}


