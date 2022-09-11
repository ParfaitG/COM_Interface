Option Explicit

Dim projectHome, csvData, dbFile
Dim wShell, FSO, accApp, accDB, tbl


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
    
    ' CLEAN OUT TABLE
    Set accDB = accApp.CurrentDb()
    For Each tbl in accDB.TableDefs
        If tbl.Name = "metals" Then
            accDB.Execute("DELETE FROM metals")
        End if
    Next
    
    ' IMPORT CSV
    Call accApp.DoCmd.TransferText(0, , "metals", csvData, True)

    'SHOW BACKGROUND APP
    accApp.UserControl = True
    Call accApp.DoCmd.OpenTable("metals")
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
    Set accDB = Nothing
    Set accApp = Nothing    
End sub


Call main