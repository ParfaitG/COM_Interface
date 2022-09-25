Option Explicit

Dim projectHome, csvData, xlFile
Dim wShell, xlApp, xlWbk, xlWks, xlQt, xlCells
Dim pvtWks, pvtCache, pvtTable

'##############################
'### PATHS
'##############################

Set wShell = WScript.CreateObject("WScript.Shell")

projectHome = wShell.ExpandEnvironmentStrings("%PROJECT_HOME%")
csvData = projectHome & "\VBS\Data\Precious_Metals_Prices.csv"
xlFile = projectHome & "\VBS\Outputs\VB_MSO_Spreadsheet.xlsx"

Set wShell = Nothing


'##############################
'### CREATE EXCEL WORKBOOK
'##############################
Sub build_wb()
    ' INITIALIZE COM OBJECT
    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False

    ' CREATE WORKBOOK
    Set xlWbk = xlApp.Workbooks.Add()

    Set xlWks = xlWbk.Worksheets(1)
    xlWks.Name = "METALS"

    ' IMPORT CSV DATA
    Set xlQt = xlWks.QueryTables.Add("TEXT;" & csvData, xlWks.Range("A1"))

    xlQt.TextFileParseType = 1
    xlQt.TextFileCommaDelimiter = True
    Call xlQt.Refresh(False)
    Call xlQt.Delete()

    ' ADJUST DEFAULT FONT
    Set xlCells = xlWks.Cells
    xlCells.Font.Name = "Arial"
    xlCells.Font.Size = "10"
    xlCells.Font.Color = 0

    'CREATE PIVOT TABLE
    Set pvtWks = xlWbk.Worksheets.Add(,xlWks)
    pvtWks.Name = "PIVOT"

    Set pvtCache = xlWbk.PivotCaches.Create(1, xlWks.UsedRange)
    Set pvtTable = pvtCache.CreatePivotTable(pvtWks.Cells(4, 2), "MetalsPivot")

    With pvtTable
        With .PivotFields("metal")
          .Orientation = 1
          .Position = 1
        End With

        .AddDataField .PivotFields("avg_price"), "Min Price", -4139
        .AddDataField .PivotFields("avg_price"), "Avg Price", -4106
        .AddDataField .PivotFields("avg_price"), "Max Price", -4136
        .AddDataField .PivotFields("avg_price"), "Std Price", -4155
    End With

    With xlApp.Union( _
        pvtWks.Columns(3), _
        pvtWks.Columns(4), _
        pvtWks.Columns(5), _
        pvtWks.Columns(6), _
        pvtWks.Columns(7) _
    )
        .NumberFormat = "$#,##0.00"
        .HorizontalAlignment = -4152
    End With

    ' ADJUST DEFAULT FONT
    Set xlCells = pvtWks.Cells
    xlCells.Font.Name = "Arial"
    xlCells.Font.Size = "10"
    xlCells.Font.Color = 0

    ' SAVE WORKBOOK
    xlWbk.SaveAs xlFile, 51

    ' SHOW BACKGROUND APP
    xlApp.Visible = True
End sub


Sub main()
    On Error Resume Next

    Call build_wb

    If Err.Number <> 0 Then
        Wscript.Echo Err.Number & ": " & Err.Description
        Err.Clear

        ' CLOSE OBJECTS
        If IsObject(xlQt) Then
            Call xlQt.Delete()
        End If
        If IsObject(xlWbk) Then
            Call xlWbk.Close(False)
        End If
        If IsObject(xlApp) Then
            xlApp.Quit()
        End If
    End If

    On Error Goto 0

    Set xlQt = Nothing
    Set xlCells = Nothing
    Set xlWks = Nothing
    Set xlWbk = Nothing
    Set xlApp = Nothing
End sub


Call main
