
library(RDCOMClient)

##############################
### PATHS
##############################
project_home <- Sys.getenv("PROJECT_HOME")
csvData <- file.path(
    project_home, "R", "Data", "Precious_Metals_Prices.csv", fsep="\\"
)
wdFile <- file.path(
    project_home, "R", "Outputs", "R_MSO_Document.docx", fsep="\\"
)
boxplotImg <- file.path(
    project_home, "R", "Data", "Precious_Metals_BoxPlot.png", fsep="\\"
)
yearplotImg <- file.path(
    project_home, "R", "Data", "Precious_Metals_YearPlot.png", fsep="\\"
)


##############################
### DATA
##############################
metals_df <- read.csv(csvData)

agg_df <- do.call(
    data.frame, 
    aggregate(
        avg_price ~ metal, metals_df, 
        function(x) c(
            min=min(x), 
            median=median(x), 
            mean=mean(x), 
            max=max(x),
            sd=sd(x)
        )
    )
)
        
agg_df[,2:6] <- round(agg_df[,2:6],4)
colnames(agg_df) <- gsub("avg_price.", "", colnames(agg_df), fixed=TRUE)


##############################
### CREATE WORD DOCUMENT
##############################
tryCatch({
    # INITIALIZE COM OBJECT
    wdApp <- COMCreate("Word.Application")
    wdApp[["DisplayAlerts"]] <- FALSE

    # CREATE DOCUMENT
    wdDoc <- wdApp$Documents()$Add()
    wdDoc[["Content"]][["Font"]][["Name"]] = "Arial"

    # ADD PARAGRAPH TITLE
    wdDoc$Paragraphs()$Add()
    wdDoc$Paragraphs(1)$Range()$InsertAfter("Precious Metals Aggregate Summary")
    wdDoc$Paragraphs()$Add()

    wdRange <- wdDoc$Content() 
    wdRange$Collapse(Direction=0)

    # ADD TABLE
    wdDoc$Tables()$Add(Range=wdRange, NumRows=5, NumColumns=6)
    wdTbl <- wdDoc$Tables(1)
    wdTbl[["Style"]] <- "Plain Table 1"

    # COLUMNS
    for(j in 1:6) {
        wdCell <- wdTbl$Cell(1,j)$Range()
        wdCell$InsertAfter(colnames(agg_df)[j])
        wdCell[["ParagraphFormat"]][["Alignment"]] <- 1
    }

    # ROWS
    for(i in 2:5) {
      for(j in 1:6) {
        wdCell <- wdTbl$Cell(i,j)$Range()
        wdCell$InsertAfter(as.character(agg_df[i-1, j]))
        if(j > 1) {            
            wdCell[["ParagraphFormat"]][["Alignment"]] <- 2
        }
      } 
    }

    # ADD PLOT IMAGES
    wdDoc$Paragraphs()$Add()
    wdDoc$Content()$InsertAfter("Precious Metals Aggregate Plots")

    wdDoc$Characters()$Last()$Select()
    wdApp$Selection()$Collapse()

    wdApp$Selection()$InlineShapes()$AddPicture(
        FileName = boxplotImg, 
        LinkToFile = FALSE,
        SaveWithDocument = TRUE
    )

    wdDoc$Paragraphs()$Add()

    wdDoc$Characters()$Last()$Select()
    wdApp$Selection()$Collapse()

    wdApp$Selection()$InlineShapes()$AddPicture(
       FileName = yearplotImg, 
       LinkToFile = FALSE,
       SaveWithDocument = TRUE
    )
    
    # SAVE DOCUMENT
    wdDoc$SaveAs(wdFile)  
    
    # SHOW BACKGROUND APP
    wdApp[["Visible"]] <- TRUE

}, warning = identity

, error = function(e) {
    identity(e)
    # CLOSE OBJECTS
    if(exists("wdDoc")) wdDoc$Close(FALSE)
    if(exists("wdApp")) wdApp$Quit()

}, finally = {
    # RELEASE RESOURCES
    wdRange <- wdPara <- wdCell <- wdTbl <- wdDoc <- wdApp <- NULL
    rm(wdRange, wdPara, wdCell, wdTbl, wdDoc, wdApp) 
})


