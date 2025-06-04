import os

root_path = "/Users/ypfeng/OneDrive/张亚的twodrive/废止制度"
subdirs = os.listdir(root_path)

for subdir in subdirs:
    if not os.path.isdir(os.path.join(root_path, subdir)):
        continue
    subsubdirs = os.listdir(os.path.join(root_path, subdir))
    for subsubdir in subsubdirs:
        if not os.path.isdir(os.path.join(root_path, subdir, subsubdir)):
            continue
        files = os.listdir(os.path.join(root_path, subdir, subsubdir))
        pdf_files = [file for file in files if file.endswith(".pdf")]
        # if len(pdf_files) == 0:
        #     print(f"{subdir} {subsubdir} has no pdf files")
        if len(pdf_files) == 1:
            if pdf_files[0][:-4] != subsubdir:
                print(f"{subdir} {subsubdir} {pdf_files[0]}")
                # os.rename(os.path.join(root_path, subdir, subsubdir, pdf_files[0]), os.path.join(root_path, subdir, subsubdir, subsubdir + ".pdf"))