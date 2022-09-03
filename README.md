# COM_Interface
Microsoft Office COM Interface implementation and automation in R, Python, and PowerShell

Using Windows' [**Component Object Model (COM)**](https://docs.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal), scripts automate the creation and content generation of an Excel workbook, Word document, PowerPoint presentation, Access database, and Outlook email. Data includes the yearly average spot prices of precious metals (gold, silver, platinum, and palladium) from 1950-2020.

## Specific COM libraries:
- **PowerShell**: [`NewObject -ComObject`](https://docs.microsoft.com/en-us/powershell/scripting/samples/creating-.net-and-com-objects--new-object-?view=powershell-7.2#creating-com-objects-with-new-object) _(built-in cmdlet)_
- **Python**: [`pywin32`](https://github.com/mhammond/pywin32) _(for win32com.client package)_
- **R**: [`RDCOMClient`](https://www.omegahat.net/RDCOMClient/) _(from omegahat repository)_

<br/>
<br/>

<div style="text-align:center"><img src="https://github.com/ParfaitG/COM_Interface/blob/main/R/Data/R_MSO_Screenshot.png" width="800px" alt="R MS Ofice COM Automation Screenshot"/></div>
