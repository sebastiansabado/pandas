import os
import pandas as pd
import openpyxl
from openpyxl.styles import Protection
from datetime import datetime


date = datetime.now().strftime("%m_%d_%Y")

print("hello")
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
format_columns = ['Site Order ID','SKU','ChannelAdvisor Order ID','Tracking ID','Merchant SKU','Return quantity','Order quantity','UPC', 'Inventory Number']
output_path = './'
for index, dataset in enumerate(dfs_report):
    dataset = dataset.drop(columns=[col for col in dataset if col not in format_columns])



##get list of all dfs
# alldfs = [var for var in dir() if isinstance(eval(var), pd.core.frame.DataFrame)]

df_orders_export = df_orders_export.rename(columns={"Site Order ID": "Order ID"})
df_items_export = df_items_export.rename(columns={"Inventory Number": "SKU"})


df_orders_export.to_excel(r".\\df_orders_export" + date +".xlsx", index=False, encoding='utf-8',sheet_name='Report')
df_items_export.to_excel(r".\\df_items_export" + date +".xlsx", index=False, encoding='utf-8',sheet_name='Report')
df_returns_export.to_excel(r".\\df_returns_export" + date +".xlsx", index=False, encoding='utf-8',sheet_name='Report')

#combining multi-item orders
# df_orders_export.groupby('Order ID')['Merchant SKU'].agg(' '.join).reset_index()
# df_orders_export['Merchant SKU'] = str_join(df_orders_export,',','Merchant SKU')
# df_orders_export = df_orders_export.drop_duplicates(subset = ['Order ID'])



# inner_join = pd.merge(
#     df_returns_export,
#     df_orders_export,
#     on='Order ID',
#     how='inner'
# )

# final_report = pd.merge(
#     inner_join,
#     df_items_export,
#     on='SKU',

#     how='inner'
# )


# final_report.groupby('Tracking ID')['Merchant SKU','UPC'].agg(' '.join).reset_index()
# final_report['full_order'] = str_join(final_report,',','Merchant SKU','UPC')
# cols = ['Tracking ID'] + [col for col in final_report if col != 'Tracking ID']
# final_report = final_report[cols].drop_duplicates(subset = ['Tracking ID'])
# final_report.to_excel(r".\\testing_final" + date +".xlsx", index=False, encoding='utf-8',sheet_name='Report')

