
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

    accDB = accApp.CurrentDb()

    # CLEAN OUT TABLE
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

    # CREATE QUERY
    for qry in accDB.QueryDefs:
        if qry.Name == "metals_agg":
            accDB.Execute("DROP TABLE metals_agg")

    qry = accDB.CreateQueryDef(
        "metals_agg",
        (
            "SELECT metal, "
            "       Ccur(MIN(avg_price)) AS MinPrice, "
            "       Ccur(AVG(avg_price)) AS AvgPrice, "
            "       Ccur(MAX(avg_price)) AS MaxPrice, "
            "       Ccur(STDEV(avg_price)) AS StdPrice "
            "FROM metals "
            "GROUP BY metal"
        )
    )

    #SHOW BACKGROUND APP
    accApp.UserControl = True
    output = accApp.DoCmd.OpenTable("metals")
    output = accApp.DoCmd.OpenQuery("metals_agg")

except Exception as e:
    print(e)

    # CLOSE OBJECTS
    if accApp is not None:
        accApp.DoCmd.CloseDatabase()
        accApp.Quit()

finally:
    # RELEASE RESOURCES
    qry = None; tbl = None; accDB = None; accApp = None
    del tbl, accDB, accApp
