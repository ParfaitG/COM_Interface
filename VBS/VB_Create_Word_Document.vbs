Option Explicit

Dim wShell, projectHome, aggData, wdFile, boxplotImg, yearplotImg, i, j
Dim fso, inputFile, aggArr(5)
Dim wdApp, wdDoc, wdRng, wdTbl


'##############################
'### PATHS
'##############################

Set wShell = WScript.CreateObject("WScript.Shell")

projectHome = wShell.ExpandEnvironmentStrings("%PROJECT_HOME%")
aggData = projectHome & "\VBS\Data\Precious_Metals_Prices_Summary.csv"
wdFile = projectHome & "\VBS\Outputs\VB_MSO_Document.docx"
boxplotImg = projectHome & "\VBS\Data\Precious_Metals_BoxPlot.png"
yearplotImg = projectHome & "\VBS\Data\Precious_Metals_YearPlot.png"

Set wShell = Nothing


'##############################
'### DATA
'##############################

Set fso = CreateObject("Scripting.FileSystemObject")
Set inputFile = fso.OpenTextFile(aggData, 1)

i = 0
Do While inputFile.AtEndOfStream <> True
    aggArr(i) = Split(inputFile.ReadLine, ",")
    i = i + 1
Loop

Set inputFile = Nothing
Set fso = Nothing


'##############################
'### CREATE WORD DOCUMENT
'##############################
Sub build_doc() 
    ' INITIALIZE COM OBJECT
    Set wdApp = CreateObject("Word.Application")
    wdApp.DisplayAlerts = False

    ' CREATE DOCUMENT
    Set wdDoc = wdApp.Documents.Add()
    wdDoc.Content.Font.Name = "Arial"
    
    ' ADD PARAGRAPH TITLE
    Call wdDoc.Paragraphs.Add()
    Call wdDoc.Paragraphs(1).Range.InsertAfter( _
       "Precious Metals Aggregate Summary" _
    )
    Call wdDoc.Paragraphs.Add()
    Call wdDoc.Paragraphs.Add()

    Set wdRng = wdDoc.Content
    Call wdRng.Collapse(0)

    ' ADD TABLE
    Call wdDoc.Tables.Add(wdRng, 5, 6)
    Set wdTbl = wdDoc.Tables(1)
    wdTbl.Style = "Plain Table 1"

    ' COLUMNS
    For j = 1 To 6
        wdTbl.Cell(1,j).Range.InsertAfter(aggArr(0)(j-1))
        wdTbl.Cell(1,j).Range.ParagraphFormat.Alignment = 1
    Next

    ' ROWS
    For i = 2 To 5
        For j = 1 To 6
            wdTbl.Cell(i,j).Range.InsertAfter(aggArr(i-1)(j-1))
            If j > 1 Then
                wdTbl.Cell(i,j).Range.ParagraphFormat.Alignment = 2
            End If
        Next
    Next

    ' ADD PLOT IMAGES
    Call wdDoc.Paragraphs.Add()
    Call wdDoc.Content.InsertAfter("Precious Metals Aggregate Plots")
    Call wdDoc.Paragraphs.Add()
    
    Call wdDoc.Characters.Last.Select()
    Call wdApp.Selection.Collapse()
    Call wdApp.Selection.InlineShapes.AddPicture(boxplotImg, False, True)

    Call wdDoc.Paragraphs.Add()

    Call wdDoc.Characters.Last.Select()
    Call wdApp.Selection.Collapse()
    Call wdApp.Selection.InlineShapes.AddPicture(yearplotImg, False, True)

    ' SAVE DOCUMENT
    Call wdDoc.SaveAs(wdFile)

    ' SHOW BACKGROUND APP
    wdApp.Visible = True
End Sub


Sub main()
    On Error Resume Next
    
    Call build_doc
    
    If Err.Number <> 0 Then
        Wscript.Echo Err.Number & ": " & Err.Description
        Err.Clear
        
        ' CLOSE OBJECTS
        If IsObject(wdDoc) Then
            wdDoc.Close(False)
        End If
        If IsObject(wdApp) Then
            wdApp.Quit()
        End If            
    End If

    On Error Goto 0
    
    Set wdRng = Nothing
    Set wdTbl = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End sub


Call main