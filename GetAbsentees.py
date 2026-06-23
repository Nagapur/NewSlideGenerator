import pandas as pd
excel_file = '2025-07-10 CDCS-Masterlist.xlsx'
excel_file2 = '2025-07-10 CDCS-Absent.xlsx'
# Load the Excel file with multiple sheets
xls = pd.ExcelFile(excel_file)
xls2 = pd.ExcelFile(excel_file2)

#load excel and return a combined dataframe with all sheets
def loadExcelToCompare(excelFileName):
    # Read the sheet into a DataFrame
    df = pd.read_excel(excelFileName, sheet_name=None, engine='openpyxl')
    combined_df = pd.concat(df.values(), ignore_index=True)

    return combined_df

#Get dataframe that use to do comparison
fullDf = loadExcelToCompare(excel_file)
absentDf =  loadExcelToCompare(excel_file2)
#find the list of absentees by comparing master list vs actual attendance using 'Name'
missing_names = fullDf.loc[
    ~fullDf['Name'].str.strip().str.lower().isin(
        absentDf['Name'].str.strip().str.lower()
    ),
]

#Create absentee sheet in excel
with pd.ExcelWriter("2025-07-10 CDCS-Absent.xlsx", engine="openpyxl", mode="a") as writer:
    missing_names.to_excel(writer, sheet_name="MSS_Absent", index=False)

