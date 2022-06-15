from tokenize import group
import tabula
import pandas as pd
import openpyxl

directory = "//in-fileserver/Business_Office/Cash Manager/Daily Deposit Reports/DM3030/2022/Python Reports/"
user_file_name="AAL_DM3030 May Consolidated"  # change this as needed
path_with_pdf = directory + user_file_name + ".pdf"
excel_output_path = directory + user_file_name + ".xlsx"

# Read in pdf as dataframe, clean up dataframe
df = tabula.io.read_pdf(input_path=path_with_pdf, pages="all", multiple_tables=False, columns=[240, 310, 410, 480, 550, 670, 750, 840, 930], guess=False)
df = df[0]
df = df.rename(columns={
  df.columns[0]: 'Name_Description',
  df.columns[1]: 'Total',
  df.columns[2]: 'U.S._Total',
  df.columns[3]: 'CashDepositsTotal1',
  df.columns[4]: 'CashDepositsTotal2',
  df.columns[5]: 'Canada_Total'})

df = df.drop(columns={
  df.columns[6],
  df.columns[7],
  df.columns[8]})

# Filter out unnecessary rows
df = df[df["Name_Description"].str.contains('[A-Za-z]', na=True)]
df = df[~df["Name_Description"].isin([
  "Continuity Collection Agency Adjust","Continuity Collection Agency Payment","Continuity Collection Agency Recharge",'Net difference','Continuities','PayPal-Subscription','Net difference','PayPal-Specialized Service Modules','PayPal-Continuity','Customer Service Adjustment','3rd party specialized service modules','3rd party continuity','3rd party catalog','3rd party subscription - PDS','3rd party subscription','Prd Alt Account Paid','Cat Alt Account Paid','Sub Alt Account Paid','Misc Income','Other agency cash deposited today','Less V/MC/Disc','Less cash from previous days','Add unapplied cash','Difference','Grand total','Miscellaneous credits','Product fulfillment payments','Catalog sales','Total cash - Specialized Servi','Knit & Crochet Now! All Access',"Annie's Creative Studio", "*** Series Credits ***", 'Agency credit','Agency overpayment','Publications','Specialized Service Modules',"Annie's Crochet SpecialsDTP","PDS","Agency","Agency Adjustment","Total","Just CrossStitchDTP","Digital Just CrossStitchDTP","Digital Crochet SpecialsDTP","PDS","Total cash - PublicationsDTP","Series total", "Cash deposits"])]

df = df[~df["U.S._Total"].str.contains('[A-Za-z]', na=False)]
df = df[~df["Canada_Total"].isin([": Carol Green DE"])]
df = df.reset_index()
df = df.drop(columns=["index"])

###################################################################
# CALCULATE ACTUAL DEPOSIT AMOUNT BEFORE FILTERING DATAFRAME
indexes = df.index[df['Name_Description'] == 'Total cash - Continuities'].tolist()
for i in range(len(indexes)):
  indexes[i] = indexes[i] + 1
actual_deposit = df.loc[indexes]
actual_deposit = actual_deposit.reset_index().drop(columns=["index", "Total", "U.S._Total"])
numeric_columns = ["CashDepositsTotal1", "CashDepositsTotal2"]
for i in numeric_columns:
  actual_deposit[i] = actual_deposit[i].astype(str).str.replace(',', '').astype(float)
actual_deposit = actual_deposit.groupby(["Canada_Total"]).sum().reset_index()
actual_deposit['Total'] = (actual_deposit["CashDepositsTotal1"] + actual_deposit["CashDepositsTotal2"]) * 0.03 # Calculate 3% of actual deposit for continuities
actual_deposit = actual_deposit.drop(['CashDepositsTotal1', "CashDepositsTotal2"], axis=1)
actual_deposit = actual_deposit.rename(columns = {"Canada_Total" : "Name_Description"})
####################################################################

# Continue cleaning/converting main dataframe
df = df.drop(["CashDepositsTotal1", "CashDepositsTotal2", "Canada_Total"], axis=1)
df["Total"] = df["Total"].astype(str).str.replace(',', '').astype(float)

grouped_df = df.groupby(['Name_Description']).sum()
grouped_df = grouped_df.reset_index()

####################################################################
# Create dataframe for reduction rows
series_total = grouped_df[~grouped_df["Name_Description"].isin(["Series total", "Total cash - Continuities"])].sum().reset_index().drop([0], axis=0)
series_total = series_total.rename(columns = {"index" : "Name_Description"})
series_total["Total"] = series_total[0] * 0.054
series_total = series_total.drop([0], axis=1)
credit_to_inv = grouped_df[grouped_df["Name_Description"].isin(["Credit Applied to Inv."])]
reductions = pd.concat([series_total, actual_deposit])
reductions = pd.concat([reductions, credit_to_inv])
reductions = reductions.reset_index().drop(["index"], axis=1)
reductions.at[0, "Name_Description"] = "Cont. total times 5.4%"
reductions.at[1, "Name_Description"] = "Actual deposit times 3%"
####################################################################

# Combine reductions df with grouped df
grouped_df = grouped_df[~grouped_df["Name_Description"].isin(["Credit Applied to Inv.", 'Total cash - Continuities'])]
grouped_df = pd.concat([grouped_df, reductions])
grouped_df["TYPE"] = pd.NaT

grouped_df.to_excel(excel_output_path)

wb = openpyxl.load_workbook(filename=excel_output_path)
worksheet = wb.active
worksheet['D2'] = "=VLOOKUP(B2,'[DM3030 Vlookup sheet.xlsx]Sheet1'!$A:$B,2,FALSE)"  #copy of Vlookup sheet is needed for this step
wb.save(filename=excel_output_path)