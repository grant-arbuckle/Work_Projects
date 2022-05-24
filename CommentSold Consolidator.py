### this program left joins two reports ###

import pandas as pd

df1 = pd.read_csv('C:/Users/Grant_Arbuckle/Desktop/Python/Inputs/Report 11.csv')
df2 = pd.read_csv('C:/Users/Grant_Arbuckle/Desktop/Python/Inputs/Report 12.csv')
combined_df = pd.merge(df1, df2, how = 'left', on = 'Order #')
combined_df = combined_df.rename(columns={"Customer Name": "customer_name"})

#combined_pivot = pd.pivot_table(combined_df, index=['customer_name'], columns=['Sku'], aggfunc='count')
#print(combined_pivot.head())
file_date = input('ENTER DATE IN FORMAT "MM.DD.YY":')
out_path = "C:/Users/Grant_Arbuckle/Desktop/Python/Outputs/SFG Retail Order File REPORT MERGE " + file_date
writer = pd.ExcelWriter(out_path + ".xlsx", engine='xlsxwriter')
combined_df.to_excel(writer, sheet_name="Data")
writer.save()

