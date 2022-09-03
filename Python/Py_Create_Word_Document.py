
import os
import pandas as pd
import win32com.client as win32

##############################
### PATHS
##############################
project_home = os.environ["PROJECT_HOME"]
csvData = os.path.join(
    project_home, "Python", "Data", "Precious_Metals_Prices.csv"
)
wdFile = os.path.join(
    project_home, "Python", "Outputs", "Py_MSO_Document.docx"
)
boxplotImg = os.path.join(
    project_home, "Python", "Data", "Precious_Metals_BoxPlot.png"
)
yearplotImg = os.path.join(
    project_home, "Python", "Data", "Precious_Metals_YearPlot.png"
)


##############################
### DATA
##############################
metals_df = pd.read_csv(csvData)

agg_df = (
    metals_df.groupby(["metal"])["avg_price"]
        .agg(["min", "median", "mean", "max", "std"])
        .reset_index()
)

agg_df.iloc[:,1:6] = agg_df.iloc[:,1:6].round(4)


##############################
### CREATE WORD DOCUMENT
##############################
try:
    # INITIALIZE COM OBJECT
    wdApp = win32.gencache.EnsureDispatch("Word.Application")
    wdApp.DisplayAlerts = False
    
    # CREATE DOCUMENT
    wdDoc = wdApp.Documents.Add()
    wdDoc.Content.Font.Name = "Arial"
    
    # ADD PARAGRAPH TITLE
    wdDoc.Paragraphs.Add()
    wdDoc.Paragraphs(1).Range.InsertAfter("Precious Metals Aggregate Summary")
    wdDoc.Paragraphs.Add()
    wdDoc.Paragraphs.Add()

    wdRng = wdDoc.Content
    wdRng.Collapse(Direction=0)

    # ADD TABLE
    wdDoc.Tables.Add(wdRng, 5, 6)
    wdTbl = wdDoc.Tables(1)
    wdTbl.Style = "Plain Table 1"

    # COLUMNS
    for j in range(1,7):
        wdTbl.Cell(1,j).Range.InsertAfter(agg_df.columns[j-1])
        wdTbl.Cell(1,j).Range.ParagraphFormat.Alignment = 1

    # ROWS
    for i in range(2,6):
      for j in range(1,7):
        wdTbl.Cell(i,j).Range.InsertAfter(str(agg_df.iloc[i-2, j-1]))
        if j > 1:
            wdTbl.Cell(i,j).Range.ParagraphFormat.Alignment = 2

    # ADD PLOT IMAGES
    wdDoc.Paragraphs.Add()
    wdDoc.Content.InsertAfter("Precious Metals Aggregate Plots")

    wdDoc.Characters.Last.Select()
    wdApp.Selection.Collapse()
    wdApp.Selection.InlineShapes.AddPicture(boxplotImg, False, True)

    wdDoc.Paragraphs.Add()

    wdDoc.Characters.Last.Select()
    wdApp.Selection.Collapse()
    wdApp.Selection.InlineShapes.AddPicture(yearplotImg, False, True)

    # SAVE DOCUMENT
    wdDoc.SaveAs(wdFile)

    # SHOW BACKGROUND APP
    wdApp.Visible = True
    
except Exception as e:
    print(e)
    # CLOSE OBJECTS
    if wdDoc is not None:
        wdDoc.Close(False)
    if wdApp is not None:
        wdApp.Quit()
finally:
    # RELEASE RESOURCES
    wdTbl = None; wdRng = None; wdDoc = None; wdApp = None
    del wdTbl, wdRng, wdDoc, wdApp
