Option Explicit

Dim wShell, projectHome, aggData, pptFile, boxplotImg, yearplotImg, i, j
Dim fso, inputFile, aggArr(5), msoFiles, msoFile, tempDir, f
Dim olApp, olMail, olTbl, olAttachments, insp, signature


'##############################
'### PATHS
'##############################

Set wShell = WScript.CreateObject("WScript.Shell")

projectHome = wShell.ExpandEnvironmentStrings("%PROJECT_HOME%")
aggData = projectHome & "\VBS\Data\Precious_Metals_Prices_Summary.csv"
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

' HTML TABLE
olTbl = "<tr>"
' COLUMNS
For j = 0 To 5
    olTbl = olTbl & "<th>" & aggArr(0)(j) & "</th>"
Next
olTbl = olTbl & "</tr>"

' ROWS
For i = 1 To 4
    olTbl = olTbl & "<tr>"
    For j = 0 To 5
        olTbl = olTbl & "<td>" & aggArr(i)(j) & "</td>"
    Next
    olTbl = olTbl & "</tr>"
Next

' COPY ATTACHMENT SOURCE FILES FOR EMAIL
msoFiles = Array( _
    "VB_MSO_Spreadsheet.xlsx", "VB_MSO_Document.docx",  _
    "VB_MSO_Presentation.pptx", "VB_MSO_Database.accdb"  _
)

Set tempDir = fso.GetSpecialFolder(2)

' COPY FILES TO TEMP DIRECTORY
For Each msoFile in msoFiles
    f = projectHome & "\VBS\Outputs\" & msoFile
    Call fso.CopyFile(f, tempDir & "\" & msoFile)
Next

            
'##############################
'### CREATE OUTLOOK EMAIL
'##############################
Sub build_email()
    ' INITIALIZE COM OBJECT
    Set olApp = CreateObject("Outlook.Application")

    ' CREATE EMAIL
    Set olMail = olApp.CreateItem(0)
    
    ' ADD RECIPIENTS AND SUBJECT
    Set insp = olMail.GetInspector
    signature = olMail.HTMLBody
    olMail.Recipients.Add "pgasana@anl.gov"
    olMail.Subject = "Precious Metals Analysis"

    ' ADD ATTACHMENTS
    Set olAttachments = olMail.Attachments
    
    For Each msoFile in msoFiles
        olAttachments.Add tempDir & "\" & msoFile
    Next
    
    olAttachments.Add boxplotImg, 1, 0
    olAttachments.Add yearplotImg, 1, 0

    ' ADD BODY
    olMail.HTMLBody = ( _
      "<html>" & _
      "<head>" & _
      "<style type=""text/css"">" & _
      "   .aggtable {font-family: Arial; font-size: 12px; border-collapse: collapse;}" & _
      "   .aggtable th, .aggtable td {border: 1px solid #CCC; text-align: right; padding: 4px;}" & _
      "   .aggtable tr:nth-child(even) {background-color: #f2f2f2;}" & _
      "</style>" & _
      "</head>" & _
      "<body style=""font-family: Arial; font-size: 14px;"">" & _
      "<p>Hello CRUG useRs!<br/><br/>" & _
      "<p>Please find below our analysis of precious metals, powered by R!</p>" & _
      "<br/><div style=""text-align: center;""/>" & _
      "<table class=""aggtable"">" & _
      olTbl & _
      "</table><br/>" & _
      "<img src=""cid:Precious_Metals_BoxPlot.png"" width=""960""/><br/>" & _
      "<img src=""cid:Precious_Metals_YearPlot.png"" width=""960""/>" & _
      "</div>" & _
      "</p>" & _
      signature & "</p>" & _
      "</body>" & _
      "</html>" _
    )

    ' DISPLAY EMAIL
    Call olMail.Display
End Sub


Sub main()
    On Error Resume Next
    
    Call build_email
    
    If Err.Number <> 0 Then
        Wscript.Echo Err.Number & ": " & Err.Description
        Err.Clear
        
        ' CLOSE OBJECTS
        If IsObject(olMail) Then
            Call olMail.Close(False)
        End If
        
        If IsObject(olApp) Then
            Call olApp.Quit()
        End if        
    End If

    On Error Goto 0
    
    Set tempDir = Nothing
    Set fso = Nothing
    Set insp = Nothing
    Set olAttachments = Nothing
    Set olMail = Nothing
    Set olApp = Nothing
End sub


Call main
