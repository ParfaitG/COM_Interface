Option Explicit

Dim wShell, projectHome, aggData, pptFile, boxplotImg, yearplotImg, i, j
Dim fso, inputFile, aggArr(5)
Dim pptApp, pptPres, pptSlide, pptTbl, pptTitle, t, oSlide, oShape


'##############################
'### PATHS
'##############################

Set wShell = WScript.CreateObject("WScript.Shell")

projectHome = wShell.ExpandEnvironmentStrings("%PROJECT_HOME%")
aggData = projectHome & "\VBS\Data\Precious_Metals_Prices_Summary.csv"
pptFile = projectHome & "\VBS\Outputs\VB_MSO_Presentation.pptx"
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
'### CREATE POWER POINT
'##############################
Sub build_ppt()
    ' INITIALIZE COM OBJECT
    Set pptApp = CreateObject("PowerPoint.Application")

    ' CREATE PRESENTATION
    Set pptPres = pptApp.Presentations.Add(True)

    ' ADD TITLE SLIDE
    Set pptSlide = pptPres.Slides.Add(1, 1)
    Call pptSlide.Shapes(1).TextFrame.TextRange.InsertAfter( _
        "Precious Metals Analysis" _
    )
    Call pptSlide.Shapes(2).TextFrame.TextRange.InsertAfter( _
        "Powered by VBS" _
    )

    ' ADD TABLE SLIDE
    Set pptSlide = pptPres.Slides.Add(2, 16)
    Call pptSlide.Shapes(1).TextFrame.TextRange.InsertAfter( _
        "Precious Metals Avg Price Aggregation" _
    )
    
    Set pptTbl = pptSlide.Shapes.AddTable(5, 6)

    ' COLUMNS
    For j = 1 To 6
        Set t = pptTbl.Table.Cell(1, j).Shape.TextFrame.TextRange
        Call t.InsertAfter(aggArr(0)(j-1))
        t.Font.Name = "Arial"
        t.ParagraphFormat.Alignment = 2
    Next
    
    ' ROWS
    For i = 2 To 5
        For j = 1 To 6
            Set t = pptTbl.Table.Cell(i, j).Shape.TextFrame.TextRange
            t.InsertAfter(aggArr(i-1)(j-1))
            t.Font.Name = "Arial"
            If j > 1 Then
                t.ParagraphFormat.Alignment = 3
            End If
        Next
    Next
    
    ' ADD PLOT SLIDE
    Set pptSlide = pptPres.Slides.Add(3, 29)
    pptSlide.Shapes(1).TextFrame.TextRange.InsertAfter( _
        "Precious Metals Avg Price Plotting" _
    )
    Call pptSlide.Shapes.AddPicture(boxplotImg, False, True, 100, 100)

    Call pptSlide.Shapes.AddPicture(yearplotImg, False, True, 100, 100)

    ' ADJUST DEFAULT FONT
    For Each oSlide In pptPres.Slides
        For Each oShape In oSlide.Shapes
            If oShape.HasTextFrame Then
                oShape.TextFrame.TextRange.Font.Name = "Arial"
                oShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            End If
        Next
    Next
    
    pptPres.SaveAs(pptFile)

    ' SHOW BACKGROUND APP
    pptApp.Visible = True
End Sub


Sub main()
    On Error Resume Next
    
    Call build_ppt
    
    If Err.Number <> 0 Then
        Wscript.Echo Err.Number & ": " & Err.Description
        Err.Clear
        
        ' CLOSE OBJECTS
        If IsObject(pptApp) Then
            Call pptApp.Presentations(pptFile).Close(False)
            Call pptApp.Quit()
        End if        
    End If

    On Error Goto 0
    
    Set oShape = Nothing
    Set oSlide = Nothing
    Set t = Nothing
    Set pptTitle = Nothing
    Set pptTbl = Nothing
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
End sub


Call main