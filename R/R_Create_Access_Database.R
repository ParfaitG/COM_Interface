
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
    
    # CLEAN OUT TABLE
    accDB = accApp$CurrentDb()
    for(i in 1:accDB$TableDefs()$Count()) {
        if(accDB$TableDefs(i-1)[["Name"]] == "metals") { 
            accDB$Execute("DELETE FROM metals") 
        }
    }

    # IMPORT CSV    
    accApp$DoCmd()$TransferText(
       0,
       TableName="metals",
       FileName = csvData,
       HasFieldNames = TRUE
    )
    
    #SHOW BACKGROUND APP
    accApp[["UserControl"]] <- TRUE
    accApp[["Visible"]] <- TRUE
    output <- accApp$DoCmd()$OpenTable("metals")
    
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
    accDB <- accApp <- NULL
    rm(accDB, accApp)
})


