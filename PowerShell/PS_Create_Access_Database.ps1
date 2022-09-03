
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

    # CLEAN OUT TABLE
    $accDB = $accApp.CurrentDb
    foreach($tbl in $accDB.TableDefs) {
        if ($tbl.Name == "metals") {
            $accDB.Execute("DELETE FROM metals")
        }
    }
    
    # IMPORT CSV
    $accApp.DoCmd.TransferText(0, $null, "metals", $csvData, $true)

    #SHOW BACKGROUND APP
    $accApp.UserControl = $true
    $accApp.DoCmd.OpenTable("metals") | Out-Null
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
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($accApp) | Out-Null
}


