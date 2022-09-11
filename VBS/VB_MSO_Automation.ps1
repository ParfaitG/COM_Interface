cd "$Env:PROJECT_HOME\VBS"

& {
    echo "`nAutomation Start: $(Get-Date -format 'u')"

    echo "`nSTEP 1: VB_Create_Excel_Workbook.vbs - $(Get-Date -format 'u')"
    cscript VB_Create_Excel_Workbook.vbs //Nologo

    echo "`nSTEP 2: VB_Create_Word_Document.vbs - $(Get-Date -format 'u')"
    cscript VB_Create_Word_Document.vbs //Nologo
    
    echo "`nSTEP 3: VB_Create_PowerPoint.vbs  - $(Get-Date -format 'u')"
    cscript VB_Create_PowerPoint.vbs //Nologo

    echo "`nSTEP 4: VB_Create_Access_Database.vbs - $(Get-Date -format 'u')"
    cscript VB_Create_Access_Database.vbs //Nologo
        
    echo "`nSTEP 5: VB_Create_Outlook_Email.vbs - $(Get-Date -format 'u')"
    cscript VB_Create_Outlook_Email.vbs //Nologo
    
    echo "`nAutomation End: $(Get-Date -format 'u')"

} 3>&1 2>&1 > "Logs\VB_MSO_Automation_$(Get-Date -format 'yyyyMMdd').log"