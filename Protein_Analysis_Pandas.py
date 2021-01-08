"""
Protein Analysis:
The excel sheets contain data collected for several proteins for
days D0, D1, D2, D3 and D4 for two different cell types N18 and N3.

This program uses Pandas library to iterate over data for p<0.05 for N18 and p<0.05 N3 and,

1. Finds all proteins that decrease in Day4 compared with Day0 for both N3 and N18 cell types.
2. Finds all proteins that increase in Day4 compared with Day0 for both N3 and N18 cell types.
"""

import pandas as pd

"""
Read from excel sheets into two separate dataframes for N18 and N3 respectively.
Skipped first row inorder to clean data as column headers start from second row.
"""
df_n18 = pd.read_excel(r'200618-N18-N3-Result-Type1-send.xlsx', sheet_name=1, skiprows=[0])

df_n3 = pd.read_excel(r'200618-N18-N3-Result-Type1-send.xlsx', sheet_name=2, skiprows=[0])


# Merge the two dataframes based on all proteins that are present in both N18 and N3 cell types.
merged_ID = pd.merge(df_n18, df_n3, on=['ID'], how='inner')

"""
Create two new columns in merged dataframe showing difference in all the proteins from Day0 to Day4
for both N18 and N3 cell types
"""
merged_ID['n18_difference'] = merged_ID['Nat1 (N18) KO D4_x'] / merged_ID['Nat1 (N18) KO D0_x']
merged_ID['n3_difference'] = merged_ID['Nat1 (N3) KO D4_y'] / merged_ID['Nat1 (N3) KO D0_y']

# Create dataframes for increase/decrease in N18, increase/decrease in N3 , increase/decrease in both N18,N3 combined.
N18_Increase = merged_ID[merged_ID.n18_difference > 1]
N18_Decrease = merged_ID[merged_ID.n18_difference < 1]
N3_Increase = merged_ID[merged_ID.n3_difference > 1]
N3_Decrease = merged_ID[merged_ID.n3_difference < 1]
merged_N18N3_Inc = pd.merge(N18_Increase, N3_Increase, on=['ID'], how='inner')
merged_N18N3_Dec = pd.merge(N18_Decrease, N3_Decrease, on=['ID'], how='inner')

# Write the protein analysis report to excel sheets
writer = pd.ExcelWriter('protein_analysis_results_pandas.xlsx', engine='xlsxwriter')

merged_ID[['ID', 'Info_x', 'Nat1 (N18) KO D0_x', 'Nat1 (N18) KO D4_x', 'Info_y',
           'Nat1 (N3) KO D0_y', 'Nat1 (N3) KO D4_y']].\
    to_excel(writer, sheet_name='Merge_N18_N3')

N18_Increase[['ID', 'Info_x', 'Nat1 (N18) KO D0_x', 'Nat1 (N18) KO D4_x', 'n18_difference']].\
    to_excel(writer, sheet_name='N18_Increase', index=False)
N18_Decrease[['ID', 'Info_x', 'Nat1 (N18) KO D0_x', 'Nat1 (N18) KO D4_x', 'n18_difference']].\
    to_excel(writer, sheet_name='N18_Decrease', index=False)
N3_Increase[['ID', 'Info_y', 'Nat1 (N3) KO D0_y', 'Nat1 (N3) KO D4_y', 'n3_difference']].\
    to_excel(writer, sheet_name='N3_Increase', index=False)
N3_Decrease[['ID', 'Info_y', 'Nat1 (N3) KO D0_y', 'Nat1 (N3) KO D4_y', 'n3_difference']].\
    to_excel(writer, 'N3_Decrease', index=False)

merged_N18N3_Inc[['ID', 'Info_x_x', 'Nat1 (N18) KO D0_x_x', 'Nat1 (N18) KO D4_x_x', 'Info_y_y', 'Nat1 (N3) KO D0_y_y',
                  'Nat1 (N3) KO D4_y_y']].to_excel(writer, sheet_name='Merge_N18_N3_Inc', index=False)
merged_N18N3_Dec[['ID', 'Info_x_x', 'Nat1 (N18) KO D0_x_x', 'Nat1 (N18) KO D4_x_x', 'Info_y_y', 'Nat1 (N3) KO D0_y_y',
                  'Nat1 (N3) KO D4_y_y']].to_excel(writer, sheet_name='Merge_N18_N3_Dec', index=False)

writer.save()
