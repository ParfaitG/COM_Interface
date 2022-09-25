
library(RDCOMClient)

##############################
### PATHS
##############################
project_home <- Sys.getenv("PROJECT_HOME")
csvData <- file.path(
    project_home, "R", "Data", "Precious_Metals_Prices.csv", fsep="\\"
)
xlFile <- file.path(
   project_home, "R", "Outputs", "R_MSO_Spreadsheet.xlsx", fsep="\\"
)


##############################
### CREATE EXCEL WORKBOOK
##############################
tryCatch({
    # INITIALIZE COM OBJECT
    xlApp <- COMCreate("Excel.Application")
    xlApp[["DisplayAlerts"]] <- FALSE

    # CREATE WORKBOOK
    xlWbk <- xlApp$Workbooks()$Add()

    xlWks <- xlWbk$Worksheets(1)
    xlWks[["Name"]] <- "METALS"

    # IMPORT CSV DATA
    xlQt <- xlWks$QueryTables()$Add(
      Connection=paste0("TEXT;", csvData),
      Destination=xlWks$Range("A1")
    )

    xlQt[["TextFileParseType"]] <- 1
    xlQt[["TextFileCommaDelimiter"]] <- TRUE
    xlQt$Refresh(BackgroundQuery=FALSE)
    xlQt$Delete()

    # ADJUST DEFAULT FONT
    xlCells <- xlWks$Cells()
    xlCells[["Font"]][["Name"]] = "Arial"
    xlCells[["Font"]][["Size"]] = "10"
    xlCells[["Font"]][["Color"]] = 0
    
    # CREATE PIVOT TABLE
    pvtWks <- xlWbk$Worksheets()$Add(After=xlWks)
    pvtWks[["Name"]] <- "PIVOT"
    
    pvtCache <- xlWbk$PivotCaches()$Create(1, xlWks$UsedRange())
    pvtTable <- pvtCache$CreatePivotTable(pvtWks$Cells(4, 2), "MetalsPivot")
    
    pvtFld <- pvtTable$PivotFields("metal")
    pvtFld[["Orientation"]] <- 1
    pvtFld[["Position"]] <- 1
    
    avgFld <- pvtTable$PivotFields("avg_price")
    pvtFld <- pvtTable$AddDataField(avgFld)
    pvtFld[["Function"]] <- -4139
    pvtFld[["Caption"]] <- "Min Price"
    
    pvtFld <- pvtTable$AddDataField(avgFld)
    pvtFld[["Function"]] <- -4106
    pvtFld[["Caption"]] <- "Avg Price"
    
    pvtFld <- pvtTable$AddDataField(avgFld)
    pvtFld[["Function"]] <- -4136
    pvtFld[["Caption"]] <- "Max Price"
    
    pvtFld <- pvtTable$AddDataField(avgFld)
    pvtFld[["Function"]] <- -4155
    pvtFld[["Caption"]] <- "Std Price"

    rng <- xlApp$Union(
        pvtWks$Columns(3),
        pvtWks$Columns(4),
        pvtWks$Columns(5),
        pvtWks$Columns(6),
        pvtWks$Columns(7) 
    ) 
    rng[["NumberFormat"]] <- "$#,##0.00"
    rng[["HorizontalAlignment"]] <- -4152
    
    # ADJUST DEFAULT FONT
    xlCells <- pvtWks$Cells()
    xlCells[["Font"]][["Name"]] = "Arial"
    xlCells[["Font"]][["Size"]] = "10"
    xlCells[["Font"]][["Color"]] = 0
    
    # SAVE WORKBOOK
    xlWbk$SaveAs(xlFile, 51)

    # SHOW BACKGROUND APP
    xlApp[["Visible"]] <- TRUE

}, warning = identity

, error = function(e) {
    identity(e)
    
    # CLOSE OBJECTS
    if(exists("xlQt")) xlQt$Delete()
    if(exists("xlWbk")) xlWbk$Close(FALSE)
    if(exists("xlApp")) xlApp$Quit()

}, finally = {
    # RELEASE RESOURCES
    rng <- NULL; avgFld <- NULL
    pvtFld <- NULL; pvtTable <- NULL; pvtCache <- NULL; pvtWks <- NULL
    xlQt <- NULL; xlCells <- NULL; xlWks <- NULL; xlWbk <- NULL; xlApp <- NULL
    rm(
      rng, avgFld, pvtWks, pvtCache, pvtTable, pvtFld,
      xlQt, xlCells, xlWks, xlWbk, xlApp
    )
})
