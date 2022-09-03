
import os
import win32com.client as win32

##############################
### PATHS
##############################
project_home = os.environ["PROJECT_HOME"]
csvData = os.path.join(
    project_home, "Python", "Data", "Precious_Metals_Prices.csv"
)
xlFile = os.path.join(
   project_home, "Python", "Outputs", "Py_MSO_Spreadsheet.xlsx"
)


##############################
### CREATE EXCEL WORKBOOK
##############################
try: 
    # INITIALIZE COM OBJECT
    xlApp = None; xlWbk = None; xlWks = None; xlQt = None;
    xlApp = win32.gencache.EnsureDispatch("Excel.Application")
    xlApp.DisplayAlerts = False

    # CREATE WORKBOOK
    xlWbk = xlApp.Workbooks.Add()

    xlWks = xlWbk.Worksheets(1)
    xlWks.Name = "METALS"

    # IMPORT CSV DATA
    xlQt = xlWks.QueryTables.Add(
      Connection=f"TEXT;{csvData}",
      Destination=xlWks.Range("A1")
    )

    xlQt.TextFileParseType = 1
    xlQt.TextFileCommaDelimiter = True
    xlQt.Refresh(BackgroundQuery=False)
    xlQt.Delete()

    # ADJUST DEFAULT FONT
    xlCells = xlWks.Cells
    xlCells.Font.Name = "Arial"
    xlCells.Font.Size = "10"
    xlCells.Font.Color = 0

    # SAVE WORKBOOK
    xlWbk.SaveAs(xlFile)

    # SHOW BACKGROUND APP
    xlApp.Visible = True

except Exception as e:
    print(e)
    # CLOSE OBJECTS
    if xlQt is not None:
        xlQt.Delete()
    if xlWbk is not None:
        xlWbk.Close(False)
    if xlApp is not None:
        xlApp.Quit()

finally:
    # RELEASE RESOURCES
    xlQt = None; xlCells = None; xlWks = None; xlWbk = None; xlApp = None
    del xlQt, xlCells, xlWks, xlWbk, xlApp



