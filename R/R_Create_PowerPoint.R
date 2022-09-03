
library(RDCOMClient)

##############################
### PATHS
##############################
project_home <- Sys.getenv("PROJECT_HOME")
csvData <- file.path(
    project_home, "R", "Data", "Precious_Metals_Prices.csv", fsep="\\"
)
pptFile <- file.path(
    project_home, "R", "Outputs", "R_MSO_Presentation.pptx", fsep="\\"
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
### CREATE POWER POINT
##############################
tryCatch({
    # INITIALIZE COM OBJECT
    pptApp <- COMCreate("PowerPoint.Application")

    # CREATE PRESENTATION
    pptPres <- pptApp$Presentations()$Add(TRUE)

    # ADD TITLE SLIDE
    pptSlide <- pptPres$Slides()$Add(Index=1, Layout=1)
    pptSlide$Shapes(1)[["TextFrame"]][["TextRange"]]$InsertAfter(
        "Precious Metals Analysis"
    )
    pptSlide$Shapes(2)[["TextFrame"]][["TextRange"]]$InsertAfter(
        "Powered by R"
    )

    # ADD TABLE SLIDE
    pptSlide <- pptPres$Slides()$Add(Index=2, Layout=16)

    pptSlide$Shapes(1)[["TextFrame"]][["TextRange"]]$InsertAfter(
        "Precious Metals Avg Price Aggregation"
    )
    pptTbl <- pptSlide$Shapes()$AddTable(5, 6)

    # COLUMNS
    for(j in 1:6) {
      t <- pptTbl$Table()$Cell(1, j)$Shape()[["TextFrame"]][["TextRange"]]
      t$InsertAfter(colnames(agg_df)[j])
      t[["Font"]][["Name"]] <- "Arial"
      t[["ParagraphFormat"]][["Alignment"]] = 2
    }

    # ROWS
    for(i in 2:5) {
      for(j in 1:6) {
        t <- pptTbl$Table()$Cell(i, j)$Shape()[["TextFrame"]][["TextRange"]]
        t$InsertAfter(as.character(agg_df[i-1, j]))
        t[["Font"]][["Name"]] <- "Arial"
        if(j > 1) t[["ParagraphFormat"]][["Alignment"]] = 3
      }
    }

    # ADD PLOT SLIDE
    pptSlide <- pptPres$Slides()$Add(Index=3, Layout=29)
    pptSlide$Shapes(1)[["TextFrame"]][["TextRange"]]$InsertAfter(
        "Precious Metals Avg Price Plotting"
    )
    pptSlide$Shapes()$AddPicture(
        FileName = boxplotImg,
        LinkToFile = FALSE,
        SaveWithDocument = TRUE,
        Left = 100,
        Top = 100
    )

    pptSlide$Shapes()$AddPicture(
        FileName = yearplotImg,
        LinkToFile = FALSE,
        SaveWithDocument = TRUE,
        Left = 100,
        Top = 100
    )

    # ADJUST DEFAULT FONT
    for(i in 1:pptPres$Slides()$Count()) {
        for(j in 1:pptPres$Slides(i)$Shapes()$Count()) {
            if (pptPres$Slides(i)$Shapes(j)$HasTextFrame()) {
                t <- pptPres$Slides(i)$Shapes(j)
                t[["TextFrame"]][["TextRange"]][["Font"]][["Name"]] = "Arial"
            }
        }
    }

    pptPres$SaveAs(pptFile)

    # SHOW BACKGROUND APP
    pptApp[["Visible"]] <- TRUE

}, warning = identity

, error = function(e) {
    identity(e)
    # CLOSE OBJECTS
    if(exists("pptPres")) pptApp$Presentations(pptFile)$Close(FALSE)
    if(exists("pptApp")) pptApp$Quit()

}, finally = {
    # RELEASE RESOURCES
    pptTitle <- pptTbl <- pptSlide <- pptPres <- pptApp <- NULL
    rm(pptTitle, pptTbl, pptSlide, pptPres, pptApp)
})