

##############################
### PATHS
##############################
$project_home = $Env:PROJECT_HOME
$aggData = "$($project_home)\PowerShell\Data\Precious_Metals_Prices_Summary.csv"
$wdFile = "$($project_home)\PowerShell\Outputs\PS_MSO_Document.docx"
$boxplotImg = "$($project_home)\PowerShell\Data\Precious_Metals_BoxPlot.png"
$yearplotImg = "$($project_home)\PowerShell\Data\Precious_Metals_YearPlot.png"


##############################
### DATA
##############################

$agg_data = Import-Csv -Path $aggData 
$agg_cols = $agg_data[0].psobject.properties.name


##############################
### CREATE WORD DOCUMENT
##############################
try {
    # INITIALIZE COM OBJECT
    $wdApp = New-Object -ComObject Word.Application

    # CREATE DOCUMENT
    $wdDoc = $wdApp.Documents.Add()
    $wdDoc.Content.Font.Name = "Arial"

    # ADD PARAGRAPH TITLE
    $wdDoc.Paragraphs.Add() | Out-Null
    $wdDoc.Paragraphs(1).Range.InsertAfter(
        "Precious Metals Aggregate Summary"
    )
    $wdDoc.Paragraphs.Add() | Out-Null
    $wdDoc.Paragraphs.Add() | Out-Null

    $wdRange = $wdDoc.Content()
    $wdRange.Collapse(0) | Out-Null

    # ADD TABLE
    $output = $wdDoc.Tables.Add($wdRange, 5, 6)
    $wdTbl = $wdDoc.Tables(1)
    $output = $wdTbl.Style = "Plain Table 1"

    # COLUMNS
    for($j = 1; $j -lt 7; $j++) {
        $output = $wdTbl.Cell(1,$j).Range.InsertAfter($agg_cols[$j-1])
        $wdTbl.Cell(1,$j).Range.ParagraphFormat.Alignment = 1
    }
    
    # ROWS
    for($i = 2; $i -lt 6; $i++) {
      for($j = 1; $j -lt 7; $j++) {
        $output = $wdTbl.Cell($i,$j).Range.InsertAfter(
            $agg_data[$i-2].PSObject.Properties[$agg_cols[$j-1]].value
        )
        if($j -gt 1) {
            $wdTbl.Cell($i,$j).Range.ParagraphFormat.Alignment = 2
        }
      }
    }

    # ADD PLOT IMAGES
    $wdDoc.Paragraphs.Add() | Out-Null
    $wdDoc.Content.InsertAfter("Precious Metals Aggregate Plots") | Out-Null

    $wdDoc.Characters.Last.Select() | Out-Null
    $wdApp.Selection.Collapse() | Out-Null
    $wdApp.Selection.InlineShapes.AddPicture($boxplotImg, $false, $true) | Out-Null

    $wdDoc.Paragraphs.Add() | Out-Null

    $wdDoc.Characters.Last.Select() | Out-Null
    $wdApp.Selection.Collapse() | Out-Null
    $wdApp.Selection.InlineShapes.AddPicture($yearplotImg, $false, $true) | Out-Null

    # SAVE DOCUMENT
    $wdDoc.SaveAs($wdFile)

    # SHOW BACKGROUND APP
    $wdApp.Visible = $true
}
catch {
    Write-Host $_
    
    # CLOSE OBJECTS
    if ($wdDoc) {
        $wdDoc.Close($false)
    }
    if ($wdApp) {
        $wdApp.Quit()
    }
}
finally {
    # RELEASE RESOURCES
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdTbl) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdRange) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdDoc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdApp) | Out-Null
}

