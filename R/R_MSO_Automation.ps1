cd "$Env:PROJECT_HOME\R"

& {
    echo "`nAutomation Start: $(Get-Date -format 'u')"

    echo "`nSTEP 1: R_Create_Excel_Workbook.R - $(Get-Date -format 'u')"
    Rscript R_Create_Excel_Workbook.R

    echo "`nSTEP 2: R_Create_Word_Document.R - $(Get-Date -format 'u')"
    Rscript R_Create_Word_Document.R
    
    echo "`nSTEP 3: R_Create_PowerPoint.R  - $(Get-Date -format 'u')"
    Rscript R_Create_PowerPoint.R

    echo "`nSTEP 4: R_Create_Access_Database.R - $(Get-Date -format 'u')"
    Rscript R_Create_Access_Database.R
    
    echo "`nSTEP 5: R_Create_Outlook_Email.R - $(Get-Date -format 'u')"
    Rscript R_Create_Outlook_Email.R
    
    echo "`nAutomation End: $(Get-Date -format 'u')"

} 3>&1 2>&1 > "Logs\R_MSO_Automation_$(Get-Date -format 'yyyyMMdd').log"
