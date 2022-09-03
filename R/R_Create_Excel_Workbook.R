
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
    
    # SAVE WORKBOOK
    xlWbk$SaveAs(xlFile)

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
    xlQt <- xlCells <- xlWks <- xlWbk <- xlApp <- NULL
    rm(xlQt, xlCells, xlWks, xlWbk, xlApp)
})



