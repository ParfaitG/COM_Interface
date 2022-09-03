
library(RDCOMClient)

##############################
### PATHS
##############################
project_home <- Sys.getenv("PROJECT_HOME")
csvData <- file.path(
    project_home, "R", "Data", "Precious_Metals_Prices.csv", fsep="\\"
)
msoFiles <- c(
    "R_MSO_Spreadsheet.xlsx", "R_MSO_Document.docx", 
    "R_MSO_Presentation.pptx", "R_MSO_Database.accdb"
)

# COPY ATTACHMENT SOURCE FILES FOR EMAIL
td <- tempdir()
td <- file.path(td, fsep="\\")

tryCatch({
    # INITIALIZE COM OBJECT
    objFSO <- COMCreate("Scripting.FileSystemObject")
    
    # COPY FILES TO TEMP DIRECTORY
    for(msoFile in msoFiles) {
        f <- file.path(project_home, "R", "Outputs", msoFile, fsep="\\")
        objFSO$CopyFile(f, file.path(td, msoFile, fsep="\\"))
    }
}, warning = identity 
 , error = identity
 , finally = {
    objFSO <- NULL
    rm(objFSO)
})    
    
xlFile <- file.path(td, "R_MSO_Spreadsheet.xlsx", fsep="\\")
wdFile <- file.path(td, "R_MSO_Document.docx", fsep="\\")
pptFile <- file.path(td, "R_MSO_Presentation.pptx", fsep="\\")
accFile <- file.path(td, "R_MSO_Database.accdb", fsep="\\")

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


##############################
### CREATE OUTLOOK EMAIL
##############################
tryCatch({
    # INITIALIZE COM OBJECT
    olApp <- COMCreate("Outlook.Application")
        
    # CREATE EMAIL
    olMail <- olApp$CreateItem(0)                                         

    olHdr <- paste0(
        "<tr><th>", 
        paste(gsub("\\.", " ", names(agg_df)), collapse="</th><th>"), 
        "</th></tr>",
        collapse=""
    )
                    
    olTbl <- paste0(
        apply(
            agg_df, 1, function(row) paste0(
                "<tr><td>", paste0(row, collapse="</td><td>$"), "</td></tr>"
            )
        ),
        collapse=""
    )

    # ADD RECIPIENTS AND SUBJECT
    insp <- olMail$GetInspector()
    signature <- olMail[["HTMLBody"]]                                   
    olMail[["Recipients"]]$Add("pgasana@anl.gov")                     
    olMail[["Subject"]] <- "Precious Metals Analysis"

    # ADD ATTACHMENTS
    olAttachments <- olMail$Attachments()
    olAttachments$Add(xlFile, 1)
    olAttachments$Add(wdFile, 1)
    olAttachments$Add(pptFile, 1)
    olAttachments$Add(accFile, 1)
    olAttachments$Add(boxplotImg, 1, 0)
    olAttachments$Add(yearplotImg, 1, 0)

    # ADD BODY
    olMail[["HTMLBody"]] <- paste0(
      '<html>',
      '<head>',
      '<style type="text/css">
         .aggtable {font-family: Arial; font-size: 12px; border-collapse: collapse;}
         .aggtable th, .aggtable td {border: 1px solid #CCC; text-align: right; padding: 2px;}
       .aggtable tr:nth-child(even) {background-color: #f2f2f2;}
      </style>',
      '</head>',
      '<body style="font-family: Arial; font-size: 12px;">',
      '<p>Hello CRUG useRs!<br/><br/>', 
      '<p>Please find below our analysis of precious metals, powered by R!</p>',
      '<br/><div style="text-align: center;"/>',
      '<table class="aggtable">',
      olHdr,
      olTbl,
      '</table><br/>',
      '<img src="cid:Precious_Metals_BoxPlot.png"/><br/>',
      '<img src="cid:Precious_Metals_YearPlot.png"/>',
      '</div>',
      signature, '</p>',
      '</body></html>'
    )
    
    # DISPLAY EMAIL
    output <- olMail$Display()                                                      
    
}, warning = identity

, error = function(e) {
    identity(e)
    # CLOSE OBJECTS
    if(exists("olMail")) olMail$Close(FALSE)
    if(exists("olApp")) olApp$Quit()

}, finally = {
    # RELEASE RESOURCES
    signature <- olTbl <- olAttachments <- olMail <- olApp <- NULL
    rm(signature, olTbl, olAttachments, olMail, olApp)
})

