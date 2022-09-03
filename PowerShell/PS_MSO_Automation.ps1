cd "$Env:PROJECT_HOME\PowerShell"

& {
    echo "`nAutomation Start: $(Get-Date -format 'u')"

    echo "`nSTEP 1: PS_Create_Excel_Workbook.ps1 - $(Get-Date -format 'u')"
    PowerShell -ExecutionPolicy bypass -File PS_Create_Excel_Workbook.ps1

    echo "`nSTEP 2: PS_Create_Word_Document.ps1 - $(Get-Date -format 'u')"
    PowerShell -ExecutionPolicy bypass -File PS_Create_Word_Document.ps1
    
    echo "`nSTEP 3: PS_Create_PowerPoint.ps1  - $(Get-Date -format 'u')"
    PowerShell -ExecutionPolicy bypass -File PS_Create_PowerPoint.ps1

    echo "`nSTEP 4: PS_Create_Access_Database.ps1 - $(Get-Date -format 'u')"
    PowerShell -ExecutionPolicy bypass -File PS_Create_Access_Database.ps1
    
    echo "`nSTEP 5: PS_Create_Outlook_Email.ps1 - $(Get-Date -format 'u')"
    PowerShell -ExecutionPolicy bypass -File PS_Create_Outlook_Email.ps1
    
    echo "`nAutomation End: $(Get-Date -format 'u')"

} 3>&1 2>&1 > "Logs\PS_MSO_Automation_$(Get-Date -format 'yyyyMMdd').log"