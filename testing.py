import os
import pandas as pd

returns_files = []
merged_file = pd.DataFrame()

path = 'returns_\\'

for root, directories, files in os.walk(path, topdown=True):
    for file_name in files:
        if file_name.endswith(".tsv"):
            if root.endswith("ab"):
                continue
            else:
                returns_files.append(pd.read_csv(os.path.join(root,file_name), sep="\t",encoding= 'unicode_escape'))

for return_file in returns_files:

    merged_file = merged_file.append(return_file, ignore_index=True
    )

merged_file.to_excel('returns_report_final.xlsx',index=False)