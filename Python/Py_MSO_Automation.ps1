cd "$Env:PROJECT_HOME\Python"

& {
    echo "`nAutomation Start: $(Get-Date -format 'u')"

    echo "`nSTEP 1: Py_Create_Excel_Workbook.py - $(Get-Date -format 'u')"
    python Py_Create_Excel_Workbook.py

    echo "`nSTEP 2: Py_Create_Word_Document.py - $(Get-Date -format 'u')"
    python Py_Create_Word_Document.py
    
    echo "`nSTEP 3: Py_Create_PowerPoint.py  - $(Get-Date -format 'u')"
    python Py_Create_PowerPoint.py

    echo "`nSTEP 4: Py_Create_Access_Database.py - $(Get-Date -format 'u')"
    python Py_Create_Access_Database.py
        
    echo "`nSTEP 5: Py_Create_Outlook_Email.py - $(Get-Date -format 'u')"
    python Py_Create_Outlook_Email.py
    
    echo "`nAutomation End: $(Get-Date -format 'u')"

} 3>&1 2>&1 > "Logs\Py_MSO_Automation_$(Get-Date -format 'yyyyMMdd').log"