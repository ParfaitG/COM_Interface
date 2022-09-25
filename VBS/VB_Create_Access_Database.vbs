Option Explicit

Dim projectHome, csvData, dbFile
Dim wShell, FSO, accApp, accDB, tbl, qDef


'##############################
'### PATHS
'##############################

Set wShell = WScript.CreateObject("WScript.Shell")

projectHome = wShell.ExpandEnvironmentStrings("%PROJECT_HOME%")
csvData = projectHome & "\VBS\Data\Precious_Metals_Prices.csv"
dbFile = projectHome & "\VBS\Outputs\VB_MSO_Database.accdb"

Set wShell = Nothing


'##############################
'### CREATE ACCESS DATABASE
'##############################
Sub build_db()
    ' INITIALIZE COM OBJECT
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set accApp = CreateObject("Access.Application")

    If fso.FileExists(projectHome & "\VBS\Outputs\VB_MSO_Database.accdb") Then
        ' OPEN DATABASE
        Call accApp.OpenCurrentDatabase(dbFile)
    Else
        ' CREATE DATABASE
        Call accApp.NewCurrentDatabase(dbFile)
    End If

    Set accDB = accApp.CurrentDb()

    ' CLEAN OUT TABLE
    For Each tbl in accDB.TableDefs
        If tbl.Name = "metals" Then
            accDB.Execute("DELETE FROM metals")
        End if
    Next

    ' IMPORT CSV
    Call accApp.DoCmd.TransferText(0, , "metals", csvData, True)

    ' CREATE QUERY
    For Each qDef in accDB.QueryDefs
        If qDef.Name = "metals_agg" Then
            accDB.Execute("DROP TABLE metals_agg")
        End if
    Next

    Set qDef = accDB.CreateQueryDef("metals_agg", _
        "SELECT metal, " & _
        "       CCur(MIN(avg_price)) AS MinPrice," & _
        "       CCur(AVG(avg_price)) AS AvgPrice," & _
        "       CCur(MAX(avg_price)) AS MaxPrice," & _
        "       Ccur(STDEV(avg_price)) AS StdPrice" & _
        " FROM metals" & _
        " GROUP BY metal" _
    )
    Set qDef = Nothing

    'SHOW BACKGROUND APP
    accApp.UserControl = True
    Call accApp.DoCmd.OpenTable("metals")
    Call accApp.DoCmd.OpenQuery("metals_agg")
End sub

Sub main()
    On Error Resume Next

    Call build_db

    If Err.Number <> 0 Then
        Wscript.Echo Err.Number & ": " & Err.Description
        Err.Clear

        ' CLOSE OBJECTS
        If IsObject(accApp) Then
            Call accApp.DoCmd.CloseDatabase()
            Call accApp.Quit()
        End If
    End If

    On Error Goto 0

    Set tbl = Nothing
    Set qDef = Nothing
    Set accDB = Nothing
    Set accApp = Nothing
End sub


Call main
