
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

    # CREATE PIVOT TABLE
    pvtWks = xlWbk.Worksheets.Add(After=xlWks)
    pvtWks.Name = "PIVOT"

    pvtCache = xlWbk.PivotCaches().Create(1, xlWks.UsedRange)
    pvtTable = pvtCache.CreatePivotTable(pvtWks.Cells(4, 2), "MetalsPivot")

    pvtFld = pvtTable.PivotFields("metal")
    pvtFld.Orientation = 1
    pvtFld.Position = 1

    avgFld = pvtTable.PivotFields("avg_price")
    pvtTable.AddDataField(avgFld, "Min Price", -4139)
    pvtTable.AddDataField(avgFld, "Avg Price", -4106)
    pvtTable.AddDataField(avgFld, "Max Price", -4136)
    pvtTable.AddDataField(avgFld, "Std Price", -4155)

    rng = xlApp.Union(
        pvtWks.Columns(3),
        pvtWks.Columns(4),
        pvtWks.Columns(5),
        pvtWks.Columns(6),
        pvtWks.Columns(7)
    )
    rng.NumberFormat = "$#,##0.00"
    rng.HorizontalAlignment = -4152

    # ADJUST DEFAULT FONT
    xlCells = pvtWks.Cells
    xlCells.Font.Name = "Arial"
    xlCells.Font.Size = "10"
    xlCells.Font.Color = 0

    # SAVE WORKBOOK
    xlWbk.SaveAs(xlFile, 51)

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
    rng = None; avgFld = None
    pvtFld = None; pvtTable = None; pvtCache = None; pvtWks = None
    xlQt = None; xlCells = None; xlWks = None; xlWbk = None; xlApp = None
    del xlQt, xlCells, xlWks, xlWbk, xlApp
