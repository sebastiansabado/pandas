import os
import pandas as pd
import openpyxl
from openpyxl.styles import Protection
from datetime import datetime

date = datetime.now().strftime("%m_%d_%Y")


def str_join(df, sep, *cols):
    from functools import reduce
    return reduce(lambda x, y: x.astype(str).str.cat(y.astype(str), sep=sep),[df[col] for col in cols])

#######################################################################################################################Create function
returns_files = []
df_returns_export = pd.DataFrame()

path = 'returns_\\'


for root, directories, files in os.walk(path, topdown=True):
    for file_name in files:
        if file_name.endswith(".tsv"):
            if root.endswith("ab"):
                continue
            else:
                returns_files.append(pd.read_csv(os.path.join(root,file_name), sep="\t",encoding= 'unicode_escape'))

for return_file in returns_files:

    df_returns_export = df_returns_export.append(return_file, ignore_index=True)

################### Items Export
items_files = []
df_items_export = pd.DataFrame()

path = 'returns_\\'

for root, directories, files in os.walk(path, topdown=True):
    for file_name in files:
        if file_name.endswith(".xlsx"):
            if root.endswith("ab"):
                continue
            else:
                items_files.append(pd.read_excel(os.path.join(root,file_name)))

for item_file in items_files:

    df_items_export = df_items_export.append(item_file, ignore_index=True)

############ Orders_export
orders_files = []
df_orders_export = pd.DataFrame()

path = 'returns_\\'

for root, directories, files in os.walk(path, topdown=True):
    for file_name in files:
        if file_name.endswith(".csv"):
            if root.endswith("ab"):
                continue
            else:
                orders_files.append(pd.read_csv(os.path.join(root,file_name),encoding= 'unicode_escape'))

for order_file in orders_files:

    df_orders_export = df_orders_export.append(order_file, ignore_index=True)

#######################################################################################################################Create function

dfs_report =[df_orders_export,df_items_export,df_returns_export]
format_columns = ['Site Order ID','SKU','ChannelAdvisor Order ID','Tracking ID','Merchant SKU','Return quantity','Order quantity','UPC', 'Inventory Number','Refunded Amount','Order ID']
output_path = './'
for index, dataset in enumerate(dfs_report):
    dataset = dataset.drop(columns=[col for col in dataset if col not in format_columns], inplace=True)


df_orders_export = df_orders_export.rename(columns={"Site Order ID": "Order ID"})
df_items_export = df_items_export.rename(columns={"Inventory Number": "SKU"})
df_returns_export = df_returns_export.rename(columns={"Merchant SKU": "SKU"})



inner_join = pd.merge(
    df_orders_export,
    df_items_export,
    on='SKU'
)

inner_join = inner_join.groupby('Order ID')['SKU','UPC'].agg(','.join).reset_index()
inner_join['full_order'] = str_join(inner_join,',','SKU','UPC')

final_report = df_returns_export.merge(inner_join,how='outer')

# final_report.groupby('Tracking ID')['Merchant SKU','UPC'].agg(' '.join).reset_index()
# final_report['full_order'] = str_join(final_report,',','Merchant SKU','UPC')
cols = ['Tracking ID'] + [col for col in final_report if col != 'Tracking ID']
# final_report = final_report[cols].drop_duplicates(subset = ['Tracking ID'])
final_report = final_report[cols]
final_report.to_excel(r".\\returns_" + date +".xlsx", index=False, encoding='utf-8',sheet_name='Report')


wb = openpyxl.load_workbook(r"returns_"+ date + ".xlsx")

ws2 = wb.create_sheet(title="Scan")
                                                                                                              
scans_titles = ['Tracking Number',	'UPC/SKU',	'Status',	'Order Number',	'CA Order ID',	'Item info']
scans_columns = ['A','B','C','D','E','F']

# write headers and excel functions
for x in range(6):

    ws2[ str(scans_columns[x]) + str(1)] = scans_titles[x]
    if scans_titles[x] == 'Status':
        for i in range(2,500):

            ws2['C'+ str(i)] = '=IF(B'+str(i)+'="","",IF(ISNUMBER(SEARCH(B'+str(i)+',F'+str(i)+'))= FALSE,"WRONG RETURN",""))'
    elif scans_titles[x] =='Order Number':
        for i in range(2,500):
            ws2['D'+str(i)] = '=IFERROR(VLOOKUP(A' +str(i)+ ',Report!A:H,2,FALSE),"")'
    elif scans_titles[x] =='CA Order ID':
        for i in range(2,500):
            ws2['E'+str(i)] = '=IFERROR(VLOOKUP(A'+str(i)+',Report!A:H,4,FALSE),"")'
    elif scans_titles[x] =='Item info':
        for i in range(2,500):
            ws2['F'+str(i)] = '=IF(A'+str(i) + '="","",VLOOKUP(A'+str(i)+',Report!A:H,8,FALSE))'

# Columm formatting and protecting
ws2.protection.sheet= True

for columns in ['A','B']:
    for cell in ws2[columns]:
        cell.protection = Protection(locked=False)
        cell.number_format='@'

for columns_dim in ['A','B','C']:
    ws2.column_dimensions[columns_dim].width = 20
        

for pro_col in ['D','E','F']:
    ws2.column_dimensions[pro_col].hidden = True
    

wb['Report'].sheet_state = 'hidden'

wb.save("returns_"+date+".xlsx")