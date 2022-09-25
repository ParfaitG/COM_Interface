
library(RDCOMClient)

##############################
### PATHS
##############################
project_home <- Sys.getenv("PROJECT_HOME")
csvData <- file.path(
    project_home, "R", "Data", "Precious_Metals_Prices.csv", fsep="\\"
)
dbFile <- file.path(
    project_home, "R", "Outputs", "R_MSO_Database.accdb", fsep="\\"
)


##############################
### CREATE ACCESS DATABASE
##############################
tryCatch({
    # INITIALIZE COM OBJECT
    accApp <- COMCreate("Access.Application")
    
    if(!file.exists(dbFile)) {
        # CREATE DATABASE
        output <- accApp$NewCurrentDatabase(dbFile)
    } else {
        # OPEN DATABASE
        output <- accApp$OpenCurrentDatabase(dbFile)
    }
    
    accDB = accApp$CurrentDb()

    # CLEAN OUT TABLE
    if(accDB$TableDefs()$Count() > 0) {
        for(i in 1:accDB$TableDefs()$Count()) {
            if(accDB$TableDefs(i-1)[["Name"]] == "metals") { 
                accDB$Execute("DELETE FROM metals") 
            }
        }
    }
    # IMPORT CSV    
    accApp$DoCmd()$TransferText(
       0,
       TableName="metals",
       FileName = csvData,
       HasFieldNames = TRUE
    )
    
    # CREATE QUERY
    if(accDB$QueryDefs()$Count() > 0) {
        for(i in 1:accDB$QueryDefs()$Count()) {
            if(accDB$QueryDefs(i-1)[["Name"]] == "metals_agg") { 
                accDB$Execute("DROP TABLE metals_agg") 
            }
        }
    }
    accDB$CreateQueryDef(
        "metals_agg", 
        paste0(
            "SELECT metal, ",
            "       CCur(MIN(avg_price)) AS MinPrice,",
            "       CCur(AVG(avg_price)) AS AvgPrice,",
            "       CCur(MAX(avg_price)) AS MaxPrice,",
            "       Ccur(STDEV(avg_price)) AS StdPrice",
            " FROM metals",
            " GROUP BY metal"
        )
    )
    
    #SHOW BACKGROUND APP
    accApp[["UserControl"]] <- TRUE
    accApp[["Visible"]] <- TRUE
    output <- accApp$DoCmd()$OpenTable("metals")
    output <- accApp$DoCmd()$OpenQuery("metals_agg")
    
}, warning = identity

, error = function(e) {
    identity(e)
    # CLOSE OBJECTS
    if(exists("accApp")) {
        accApp$DoCmd()$CloseDatabase()
        accApp$Quit()
    }

}, finally = {
    # RELEASE RESOURCES
    accDB <- NULL; accApp <- NULL
    rm(accDB, accApp)
})


