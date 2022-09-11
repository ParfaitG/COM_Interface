
import os
import tempfile
import pandas as pd
import win32com.client as win32

##############################
### PATHS
##############################
project_home = os.environ["PROJECT_HOME"]
csvData = os.path.join(
    project_home, "Python", "Data", "Precious_Metals_Prices.csv"
)
boxplotImg = os.path.join(
    project_home, "Python", "Data", "Precious_Metals_BoxPlot.png"
)
yearplotImg =os.path.join(
    project_home, "Python", "Data", "Precious_Metals_YearPlot.png"
)

##############################
### DATA
##############################
metals_df = pd.read_csv(csvData)

agg_df = (
    metals_df.groupby(["metal"])["avg_price"]
        .agg(["min", "median", "mean", "max", "std"])
        .reset_index()
)

agg_df.iloc[:,1:6] = agg_df.iloc[:,1:6].round(4)


# COPY ATTACHMENT SOURCE FILES FOR EMAIL
msoFiles = (
    "Py_MSO_Spreadsheet.xlsx", "Py_MSO_Document.docx",
    "Py_MSO_Presentation.pptx", "Py_MSO_Database.accdb"
)

with tempfile.TemporaryDirectory() as td:
    try:
        # INITIALIZE COM OBJECT
        objFSO = win32.gencache.EnsureDispatch("Scripting.FileSystemObject")

        # COPY FILES TO TEMP DIRECTORY
        for msoFile in msoFiles:
            f = os.path.join(project_home, "Python", "Outputs", msoFile)
            objFSO.CopyFile(f, os.path.join(td, msoFile))

    except Exception as e:
        print(e)

    finally:
        objFSO = None
        del objFSO

    xlFile = os.path.join(td, "Py_MSO_Spreadsheet.xlsx")
    wdFile = os.path.join(td, "Py_MSO_Document.docx")
    pptFile = os.path.join(td, "Py_MSO_Presentation.pptx")
    accFile = os.path.join(td, "Py_MSO_Database.accdb")


    ##############################
    ### CREATE OUTLOOK EMAIL
    ##############################
    try:
        # INITIALIZE COM OBJECT
        olApp = None; olMail = None
        olApp = win32.gencache.EnsureDispatch("Outlook.Application")

        # CREATE EMAIL
        olMail =olApp.CreateItem(0)

        olTbl = agg_df.to_html()
        
        # ADD RECIPIENTS AND SUBJECT
        insp = olMail.GetInspector
        signature = olMail.HTMLBody
        olMail.Recipients.Add("pgasana@anl.gov")
        olMail.Subject = "Precious Metals Analysis"

        # ADD ATTACHMENTS
        olAttachments = olMail.Attachments
        olAttachments.Add(xlFile, 1)
        olAttachments.Add(wdFile, 1)
        olAttachments.Add(pptFile, 1)
        olAttachments.Add(accFile, 1)
        olAttachments.Add(boxplotImg, 1, 0)
        olAttachments.Add(yearplotImg, 1, 0)

        # ADD BODY
        olMail.HTMLBody = (
          '<html>'
          '<head>'
          '<style type="text/css">'
          '   .dataframe {font-family: Arial; font-size: 12px; border-collapse: collapse;}'
          '   .dataframe th, .dataframe td {border: 1px solid #CCC; text-align: right; padding: 4px;}'
          '   .dataframe tr:nth-child(even) {background-color: #f2f2f2;}'
          '</style>'
          '</head>'
          '<body style="font-family: Arial; font-size: 14px;">'
          '<p>Hello CRUG useRs!<br/><br/>'
          '<p>Please find below our analysis of precious metals, powered by R!</p>'
          '<br/><div style="text-align: center;"/>'
          '<table class="aggtable">'
          f'{olTbl}'
          '</table><br/>'
          '<img src="cid:Precious_Metals_BoxPlot.png" width="960"/><br/>'
          '<img src="cid:Precious_Metals_YearPlot.png" width="960"/>'
          '</div>'
          f'{signature}'
          '</p>'
          '</body>'
          '</html>'
        )

        # DISPLAY EMAIL
        output = olMail.Display()

    except Exception as e:
        print(e)
        # CLOSE OBJECTS
        if olMail is not None:
            olMail.Close(False)
        if olApp is not None:
            olApp.Quit()

    finally:
        # RELEASE RESOURCES
        insp = None; olTbl = None; olAttachments = None; 
        olMail = None; olApp = None
        del signature, olTbl, olAttachments, olMail, olApp

