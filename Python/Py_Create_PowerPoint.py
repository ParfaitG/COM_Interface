
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
pptFile = os.path.join(
    project_home, "Python", "Outputs", "Py_MSO_Presentation.pptx"
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
### CREATE POWER POINT
##############################
try:
    # INITIALIZE COM OBJECT
    pptApp = win32.gencache.EnsureDispatch("PowerPoint.Application")

    # CREATE PRESENTATION
    pptPres = pptApp.Presentations.Add(True)

    # ADD TITLE SLIDE
    pptSlide = pptPres.Slides.Add(Index=1, Layout=1)
    pptSlide.Shapes(1).TextFrame.TextRange.InsertAfter(
        "Precious Metals Analysis"
    )
    pptSlide.Shapes(2).TextFrame.TextRange.InsertAfter(
        "Powered by Python"
    )

    # ADD TABLE SLIDE
    pptSlide = pptPres.Slides.Add(Index=2, Layout=16)
    pptSlide.Shapes(1).TextFrame.TextRange.InsertAfter(
        "Precious Metals Avg Price Aggregation"
    )
    
    pptTbl = pptSlide.Shapes.AddTable(5, 6)

    # COLUMNS
    for j in range(1,7):
        t = pptTbl.Table.Cell(1, j).Shape.TextFrame.TextRange
        t.InsertAfter(agg_df.columns[j-1])
        t.Font.Name = "Arial"
        t.ParagraphFormat.Alignment = 2

    # ROWS
    for i in range(2,6):
        for j in range(1,7):
            t = pptTbl.Table.Cell(i, j).Shape.TextFrame.TextRange
            t.InsertAfter(str(agg_df.iloc[i-2, j-1]))
            t.Font.Name = "Arial"
            if j > 1:
                t.ParagraphFormat.Alignment = 3

    # ADD PLOT SLIDE
    pptSlide = pptPres.Slides.Add(Index=3, Layout=29)
    pptSlide.Shapes(1).TextFrame.TextRange.InsertAfter(
        "Precious Metals Avg Price Plotting"
    )
    pptSlide.Shapes.AddPicture(
        FileName = boxplotImg,
        LinkToFile = False,
        SaveWithDocument = True,
        Left = 100,
        Top = 100
    )

    pptSlide.Shapes.AddPicture(
        FileName = yearplotImg,
        LinkToFile = False,
        SaveWithDocument = True,
        Left = 100,
        Top = 100
    )

    # ADJUST DEFAULT FONT
    for oSlide in pptPres.Slides:
        for oShape in oSlide.Shapes:
            if oShape.HasTextFrame:
                oShape.TextFrame.TextRange.Font.Name = "Arial"
                oShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                
    pptPres.SaveAs(pptFile)

    # SHOW BACKGROUND APP
    pptApp.Visible = True

except Exception as e:
    print(e)
    # CLOSE OBJECTS        
    if pptApp is not None:
        pptApp.Presentations(pptFile).Close(False)
        pptApp.Quit()

finally:
    # RELEASE RESOURCES
    pptTitle = None; pptTbl = None; pptSlide = None; 
    pptPres = None; pptApp = None
    del pptTitle, pptTbl, pptSlide, pptPres, pptApp
