
##############################
### PATHS
##############################
$project_home = $Env:PROJECT_HOME
$aggData = "$($project_home)\PowerShell\Data\Precious_Metals_Prices_Summary.csv"
$pptFile = "$($project_home)\PowerShell\Outputs\PS_MSO_Presentation.pptx"
$boxplotImg = "$($project_home)\PowerShell\Data\Precious_Metals_BoxPlot.png"
$yearplotImg = "$($project_home)\PowerShell\Data\Precious_Metals_YearPlot.png"


##############################
### DATA
##############################

$agg_data = Import-Csv -Path $aggData 
$agg_cols = $agg_data[0].psobject.properties.name


##############################
### CREATE POWER POINT
##############################
try {
    # INITIALIZE COM OBJECT
    $pptApp = New-Object -ComObject PowerPoint.Application

    # CREATE PRESENTATION
    $pptPres = $pptApp.Presentations.Add($true)

    # ADD TITLE SLIDE
    $pptSlide = $pptPres.Slides.Add(1, 1)
    $pptSlide.Shapes(1).TextFrame.TextRange.InsertAfter(
        "Precious Metals Analysis"
    ) | Out-Null
    $pptSlide.Shapes(2).TextFrame.TextRange.InsertAfter(
        "Powered by PowerShell"
    ) | Out-Null

    # ADD TABLE SLIDE
    $pptSlide = $pptPres.Slides.Add(2, 16)

    $output = $pptSlide.Shapes(1).TextFrame.TextRange.InsertAfter(
        "Precious Metals Avg Price Aggregation"
    )
    $pptTbl = $pptSlide.Shapes.AddTable(5, 6)

    # COLUMNS
    for($j = 1; $j -lt 7; $j++) {
        $t = $pptTbl.Table.Cell(1, $j).Shape.TextFrame.TextRange
        $t.InsertAfter($agg_cols[$j-1]) | Out-Null
        $t.Font.Name = "Arial"
        $t.ParagraphFormat.Alignment = 2
    }

    # ROWS
    for($i=2; $i -lt 6; $i++){
        for($j=1; $j -lt 7; $j++){
            $t = $pptTbl.Table.Cell($i, $j).Shape.TextFrame.TextRange
            $output = $t.InsertAfter(
                $agg_data[$i-2].PSObject.Properties[$agg_cols[$j-1]].value
            )
            $t.Font.Name = "Arial"
            if($j -gt 1) {
                $t.ParagraphFormat.Alignment = 3
            }
        }
    }

    # ADD PLOT SLIDE
    $pptSlide = $pptPres.Slides.Add(3, 29)
    $output = $pptSlide.Shapes(1).TextFrame.TextRange.InsertAfter(
        "Precious Metals Avg Price Plotting"
    )
    $pptSlide.Shapes.AddPicture($boxplotImg, $false, $true, 100, 100) | Out-Null
    $pptSlide.Shapes.AddPicture($yearplotImg, $false, $true, 100, 100) | Out-Null

    # ADJUST DEFAULT FONT
    foreach($oSlide in $pptPres.Slides) {
        foreach($oShape in $oSlide.Shapes) {
            if($oShape.HasTextFrame) {
                $oShape.TextFrame.TextRange.Font.Name = "Arial"
                $oShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            }
        }
    }
    
    $pptPres.SaveAs($pptFile)

    # SHOW BACKGROUND APP
    $pptApp.Visible = 1
}
catch {
    Write-Host $_
    # CLOSE OBJECTS
    if ($pptApp) {
        $pptApp.Presentations($pptFile).Close($false)
        $pptApp.Quit()
    }
}
finally {
    # RELEASE RESOURCES
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptSlide) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptPres) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
}