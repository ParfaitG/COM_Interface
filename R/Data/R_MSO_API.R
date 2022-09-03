
library(RDCOMClient)

project_home <- file.path(
  "C:", "Users", "parfa", "OneDrive", "Documents", 
  "Databases", "Sandbox", "R_Automation"
)

##############################
### DATA
##############################
csvData <- file.path(project_home, "Precious_Metals.csv")


##############################
### CREATE EXCEL WORKBOOK
##############################
xlApp <- COMCreate("Excel.Application")
xlWbk <- xlApp$Workbooks()$Add()                                         # CREATE WORKBOOK

xlWks <- xlWbk$Worksheets(1)
xlWks[["Name"]] <- "METALS"
xlQt <- xlWks$QueryTables()$Add(Connection=paste0("TEXT;", csvData), 
                                Destination=xlWks$Range("A1"))           # IMPORT CSV DATA

xlQt[["TextFileParseType"]] <- 1
xlQt[["TextFileCommaDelimiter"]] <- TRUE
xlQt$Refresh()
xlQt$Delete()

xlWbk$SaveAs(file.path(project_home, "R_MSO_Spreadsheet.xlsx"))          # SAVE AND CLOSE WORKBOOK
xlWbk$Close(TRUE)                                                        # SAVE AND CLOSE WORKBOOK
xlApp$Quit()                                                             # CLOSE COM APP 

# RELEASE RESOURCES
xlQt <- xlWks <- xlWbk <- xlApp <- NULL
rm(xlQt, xlWks, xlWbk, xlApp)
gc()  


##############################
### CREATE WORD DOCUMENT
##############################

wdApp <- COMCreate("Word.Application")
wdDoc <- wdApp$Documents()$Add()                                         # CREATE DOCUMENT

wdDoc$Paragraphs()$Add()
wdDoc$Paragraphs(1)$Range()$InsertAfter("Precious Metals Aggregate Summary")
wdDoc$Paragraphs()$Add()
wdDoc$Paragraphs()$Add()

wdRange <- wdDoc$Content() 
wdRange$Collapse(Direction=0)

wdDoc$Tables()$Add(Range=wdRange, NumRows=6, NumColumns=6)
wdTbl <- wdDoc$Tables(1)
wdTbl[["Style"]] <- "Plain Table 1"

for(j in 1:6) {
    wdTbl$Cell(1,j)$Range()$InsertAfter(gsub("\\.", " ", names(agg_df)[j]))
}

agg_df[,2:6] <- round(agg_df[,2:6],4)

for(i in 2:6) {
  for(j in 1:6) {
    wdTbl$Cell(i,j)$Range()$InsertAfter(as.character(agg_df[i-1, j]))
  } 
}

wdDoc$Paragraphs()$Add()
wdDoc$Content()$InsertAfter("Precious Metals Aggregate Plots")

wdDoc$Characters()$Last()$Select()
wdApp$Selection()$Collapse()

wdApp$Selection()$InlineShapes()$AddPicture(FileName = paste0(path, "\\Precious_Metals_Plot.png"), 
                                            LinkToFile = FALSE,
                                            SaveWithDocument = TRUE)

wdDoc$Paragraphs()$Add()

wdDoc$Characters()$Last()$Select()
wdApp$Selection()$Collapse()

wdApp$Selection()$InlineShapes()$AddPicture(FileName = paste0(path, "\\Precious_Metals_Year.png"), 
                                            LinkToFile = FALSE,
                                            SaveWithDocument = TRUE)

wdDoc$SaveAs(paste0(path, "\\R_MSO_Document.docx"))                      # SAVE AND CLOSE DOCUMENT
wdDoc$Close(TRUE)                                                        # SAVE AND CLOSE DOCUMENT
wdApp$Quit()                                                             # QUIT COM APP 

# RELEASE RESOURCES
wdRange <- wdPara <- wdTbl <- wdDoc <- wdApp <- NULL
rm(wdRange, wdPara, wdTbl, wdDoc, wdApp)
gc()  


##############################
### CREATE POWER POINT 
##############################
pptApp <- COMCreate("PowerPoint.Application")

pptPres <- pptApp$Presentations()$Add(FALSE)                             # CREATE PRESENTATION
pptSlide <- pptPres$Slides()$Add(Index=1, Layout=1)
pptSlide$Shapes(1)[["TextFrame"]][["TextRange"]]$InsertAfter("Precious Metals Analysis")
pptSlide$Shapes(2)[["TextFrame"]][["TextRange"]]$InsertAfter("Powered by R")

pptSlide <- pptPres$Slides()$Add(Index=2, Layout=16)

pptSlide$Shapes(1)[["TextFrame"]][["TextRange"]]$InsertAfter("Precious Metals Avg Price Aggregation")
pptTbl <- pptSlide$Shapes()$AddTable(6, 6)

for(j in 1:6) {
  pptTbl$Table()$Cell(1, j)$Shape()[["TextFrame"]][["TextRange"]]$InsertAfter(gsub("\\.", " ", names(agg_df)[j]))
}

for(i in 2:6) {
  for(j in 1:6) {
    pptTbl$Table()$Cell(i, j)$Shape()[["TextFrame"]][["TextRange"]]$InsertAfter(as.character(agg_df[i-1, j]))
  } 
}

pptSlide <- pptPres$Slides()$Add(Index=3, Layout=29)
pptSlide$Shapes(1)[["TextFrame"]][["TextRange"]]$InsertAfter("Precious Metals Avg Price Plotting")
pptSlide$Shapes()$AddPicture(FileName = paste0(path, "\\Precious_Metals_Plot.png"),
                             LinkToFile = FALSE,
                             SaveWithDocument = TRUE,
                             Left = 100,
                             Top = 100)

pptSlide$Shapes()$AddPicture(FileName = paste0(path, "\\Precious_Metals_Year.png"),
                             LinkToFile = FALSE,
                             SaveWithDocument = TRUE,
                             Left = 100,
                             Top = 100)

pptPres$SaveAs(paste0(path, "\\R_MSO_Presentation.pptx"))                # SAVE AND CLOSE PRESENTATION
pptApp$Presentations(paste0(path, "\\R_MSO_Presentation.pptx"))$Close()  # CLOSE PRESENTATION
pptApp$Quit()                                                            # QUIT APP

# RELEASE RESOURCES
pptTitle <- pptTbl <- pptSlide <- pptPres <- pptApp <- NULL
rm(pptTitle, pptTbl, pptSlide, pptPres, pptApp)
gc()  


##############################
### CREATE ACCESS DATABASE
##############################
accApp <- COMCreate("Access.Application")

dbEngine <- accApp$DBEngine()
workspace <- dbEngine$Workspaces(0)

accDB <- workspace$CreateDatabase(paste0(path, "\\R_MSO_Database.accdb"),     
                                  ";LANGID=0x0409;CP=1252;COUNTRY=0", 64)       # CREATE DATABASE

accApp$OpenCurrentDatabase(paste0(path, "\\R_MSO_Database.accdb"))              # OPEN DATABASE
accApp$DoCmd()$TransferText(0, TableName="metals", FileName = csvData,
                            HasFieldNames = TRUE)                               # IMPORT CSV
accApp$DoCmd()$CloseDatabase()                                                  # CLOSE DATABASE
accApp$Quit()                                                                   # QUIT COM APP

# RELEASE RESOURCES
dbEngine <- workspace <- conn <- output <- accDB <- accApp <- NULL
rm(dbEngine, workspace, conn, output, accDB, accApp)
gc()


##############################
### CREATE OUTLOOK EMAIL
##############################

olApp <- COMCreate("Outlook.Application")
olMail <- olApp$CreateItem(0)                                         # CREATE EMAIL

olHdr <- paste0("<tr><th>", paste(gsub("\\.", " ", names(agg_df)), collapse="</th><th>"), "</th></tr>",
                collapse="")
olTbl <- paste0(apply(agg_df, 1, function(row) paste0("<tr><td>", paste0(row, collapse="</td><td>$"), "</td></tr>")),
                collapse="")

olMail$GetInspector()
signature <- olMail[["HTMLBody"]]                                   
olMail[["Recipients"]]$Add("pgasana@winston.com")                     # EDIT EMAIL CONTENT
olMail[["Subject"]] <- "some subject"

olAttachments <- olMail$Attachments()
olAttachments$Add(paste0(path, "\\R_MSO_Spreadsheet.xlsx"), 1)
olAttachments$Add(paste0(path, "\\R_MSO_Document.docx"), 1)
olAttachments$Add(paste0(path, "\\R_MSO_Presentation.pptx"), 1)
olAttachments$Add(paste0(path, "\\R_MSO_Database.accdb"), 1)
olAttachments$Add(paste0(path, "\\Precious_Metals_Plot.png"), 1, 0)
olAttachments$Add(paste0(path, "\\Precious_Metals_Year.png"), 1, 0)

olMail[["HTMLBody"]] <- paste0('<html>',
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
                               '<img src="cid:Precious_Metals_Plot.png"/><br/>',
                               '<img src="cid:Precious_Metals_Year.png"/>',
                               '</div>',
                               signature, '</p>',
                               '</body></html>')

olMail$Display()                                                      # DISPLAY EMAIL

# RELEASE RESOURCES
signature <- olTbl <- olAttachments <- olMail <- olApp <- NULL
rm(signature, olTbl, olAttachments, olMail, olApp)                   
gc()  




