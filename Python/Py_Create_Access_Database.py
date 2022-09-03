
import os
import win32com.client as win32

##############################
### PATHS
##############################
project_home = os.getenv("PROJECT_HOME")
csvData = os.path.join(
    project_home, "Python", "Data", "Precious_Metals_Prices.csv"
)
dbFile = os.path.join(
    project_home, "Python", "Outputs", "Py_MSO_Database.accdb"
)


##############################
### CREATE ACCESS DATABASE
##############################
try:
    # INITIALIZE COM OBJECT
    accApp = None
    accApp = win32.gencache.EnsureDispatch("Access.Application")

    if not os.path.exists(dbFile):
        # CREATE DATABASE
        output = accApp.NewCurrentDatabase(dbFile)
    else:
        # OPEN DATABASE
        output = accApp.OpenCurrentDatabase(dbFile)
    
    # CLEAN OUT TABLE
    accDB = accApp.CurrentDb()
    for tbl in accDB.TableDefs:
        if tbl.Name == "metals":
            accDB.Execute("DELETE FROM metals")
    
    # IMPORT CSV
    accApp.DoCmd.TransferText(
       0,
       TableName = "metals",
       FileName = csvData,
       HasFieldNames = True
    )

    #SHOW BACKGROUND APP
    accApp.UserControl = True
    output = accApp.DoCmd.OpenTable("metals")

except Exception as e:
    print(e)

    # CLOSE OBJECTS
    if accApp is not None:
        accApp.DoCmd.CloseDatabase()
        accApp.Quit()

finally:
    # RELEASE RESOURCES
    dbEngine = None; workspace = None; tbl = None; accDB = None; accApp = None
    del dbEngine, workspace, tbl, accDB, accApp



