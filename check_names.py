import pandas as pd
import os
import shutil
import subprocess


column_g = pd.read_excel('/Users/ypfeng/Downloads/废止制度目录06041348.xls', usecols="G")
# Now it's numpy array, nice
# Squeeze to 1D array and convert to list
# The real policy name we want
policy = column_g[1:].values.squeeze().tolist()
print(len(policy))

root_path = "/Users/ypfeng/OneDrive/张亚的twodrive/废止制度"
subdirs = os.listdir(root_path)

actual_names = []
for subdir in subdirs:
    if not os.path.isdir(os.path.join(root_path, subdir)):
        continue
    actual_names.append(subdir)
    files = os.listdir(os.path.join(root_path, subdir))
    if subdir in files:
        print(f"subdir name match in {subdir}: {subdir}")
        continue
    pdf_files = [file for file in files if file.endswith(".pdf")]
    if len(pdf_files) > 1:
        print(f"multiple pdf files in {subdir}: {pdf_files}")
    elif len(pdf_files) == 0:
        print(f"no pdf files in {subdir}")
    else:
        if pdf_files[0] != subdir + ".pdf":
            print(f"pdf file name mismatch in {subdir}: {pdf_files[0]}")

# print(len(actual_names))

 
# mismatched_names = set(policy) - set(actual_names)
# print(mismatched_names)

# mismatched_names = set(actual_names) - set(policy)
# print(mismatched_names)
# print(len(set(policy)))
# print(len(set(actual_names)))

# duplicates = set([x for x in actual_names if actual_names.count(x) > 1])
# print(f"duplicates in actual_names: {duplicates}")

# duplicates = set([x for x in policy if policy.count(x) > 1])
# print(f"duplicates in policy: {duplicates}")

# for p in policy:
#     if p not in actual_names:
#         print(p)

# print("--------------------------------")

# for a in actual_names:
#     if a not in policy:
#         print(a)