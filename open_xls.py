import pandas as pd
import os
import shutil
import subprocess


column_g = pd.read_excel('/Users/ypfeng/Downloads/废止制度目录06041348.xls', usecols="G")
# Now it's numpy array, nice
# Squeeze to 1D array
# The real policy name we want
policy = column_g[1:].values.squeeze()

# column_l = pd.read_excel('废止制度目录06032044.xls', usecols="L")
# index = column_l[1:].values.squeeze()

# index_to_policy = {}

# for i in range(len(policy)):
#     index_to_policy[index[i]] = policy[i]

# root_path = "/Users/yupufeng/OneDrive/张亚的twodrive/废止制度"
# subdirs = os.listdir(root_path)

# # find all unmatched subdirs
# unmatched_subdirs = []
# # old name -> new name
# matched_subdirs = {}
# found_policy = []
# for subdir in subdirs:
#     if not os.path.isdir(os.path.join(root_path, subdir)):
#         print(f"Not a directory: {subdir}")
#         continue
#     matched = False
#     for k, v in index_to_policy.items():
#         if k in subdir:
#             matched = True
#             matched_subdirs[subdir] = v
#             found_policy.append(v)
#             break
#     if not matched:
#         unmatched_subdirs.append(subdir)

# print(f"Total unmatched subdirs: {len(unmatched_subdirs)}")
# print(unmatched_subdirs)

# not_found_policy = []
# for k, v in index_to_policy.items():
#     if v not in found_policy:
#         not_found_policy.append(v)
# print(f"Total not found policy: {len(not_found_policy)}")
# print(not_found_policy)

# rename subdirs
# for old_name, new_name in matched_subdirs.items():
#     shutil.move(os.path.join(root_path, old_name), os.path.join(root_path, new_name))

# The names we know for sure are attachments
# attachments = ["通知", "通告", "履历表", "控制表", "流程图"]

# find all files in subdirs
# for subdir in subdirs:
#     if not os.path.isdir(os.path.join(root_path, subdir)):
#         print(f"Not a directory: {subdir}")
#         continue
#     os.makedirs(os.path.join(root_path, subdir, "附件"), exist_ok=True)
#     for file in os.listdir(os.path.join(root_path, subdir)):
#         if any(attachment in file for attachment in attachments):
#             shutil.move(os.path.join(root_path, subdir, file), os.path.join(root_path, subdir, "附件", file))
#             print(f"Moved {file} to {os.path.join(root_path, subdir, '附件', file)}")

# Does not work
def convert_to_pdf(input_file):
    """Convert DOC/DOCX to PDF using WPS via AppleScript"""
    try:
        # Create the AppleScript command
        output_file = input_file.rsplit('.', 1)[0] + '.pdf'
        input_file_abs = os.path.abspath(input_file)
        output_file_abs = os.path.abspath(output_file)
        output_dir = os.path.dirname(output_file_abs)
        
        applescript = f'''
        tell application "System Events"
            -- Make sure WPS is not running
            try
                tell application "wpsoffice" to quit
            end try
            delay 2
            
            -- Launch WPS and open file
            do shell script "open -a wpsoffice " & quoted form of "{input_file_abs}"
            delay 3
            
            tell process "wpsoffice"
                set frontmost to true
                delay 1
                
                -- Click File menu and Export as PDF
                click menu item "Export to PDF..." of menu "File" of menu bar 1
                delay 2
                
                -- Get the export window
                set exportWindow to window 1
                
                -- Print UI elements for debugging
                log (get entire contents of exportWindow)
                
                -- Try different methods to click the Export button
                tell exportWindow
                    -- Try by class
                    try
                        click (first button whose title is "Export")
                    on error
                        try
                            -- Try by accessibility description
                            click (first button whose description contains "Export")
                        on error
                            try
                                -- Try by role
                                click (first button whose role description is "button")
                            on error
                                try
                                    -- Try clicking the last button (Export is usually last)
                                    click last button
                                end try
                            end try
                        end try
                    end try
                end tell
                
                delay 2
                
                -- Handle the save dialog
                tell sheet 1 of window 1
                    -- Set the save location
                    set value of text field 1 to "{output_file_abs}"
                    delay 1
                    keystroke return
                    delay 2
                    
                    -- If "Allow Visit" dialog appears
                    if exists button "Allow Visit" then
                        click button "Allow Visit"
                        delay 1
                    end if
                end tell
            end tell
            
            -- Wait a bit for the export to complete
            delay 3
            
            -- Quit WPS
            tell application "wpsoffice" to quit
            delay 2
        end tell
        '''
        
        # Run the AppleScript
        process = subprocess.run(['osascript', '-e', applescript], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
        if process.returncode != 0:
            print(f"AppleScript Error: {process.stderr}")
            # Print the output to see the UI elements
            print("UI Elements:", process.stdout)
            return False
            
        # Verify if PDF was created
        if os.path.exists(output_file):
            print(f"Successfully created PDF at {output_file}")
            return True
        else:
            print(f"PDF file was not created at {output_file}")
            return False
            
    except Exception as e:
        print(f"Error converting {input_file}: {str(e)}")
        return False

# Works, but the table is not formatted well
def convert_to_pdf_libreoffice(input_file):
    """Convert DOC/DOCX to PDF using LibreOffice with specific table formatting options"""
    try:
        input_file_abs = os.path.abspath(input_file)
        output_dir = os.path.dirname(input_file_abs)
        
        # To find LibreOffice executable location:
        # $ find /Applications -name "soffice" 2>/dev/null
        # Expected output: /Applications/LibreOffice.app/Contents/MacOS/soffice
        
        # Full path to LibreOffice executable
        libreoffice_path = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
        
        # Verify LibreOffice exists
        if not os.path.exists(libreoffice_path):
            print(f"LibreOffice not found at {libreoffice_path}")
            return False
        
        # First convert to ODT format which often preserves formatting better
        cmd_to_odt = [
            libreoffice_path,
            '--headless',
            '--convert-to',
            'odt',
            '--outdir', output_dir,
            input_file_abs
        ]
        
        process = subprocess.run(cmd_to_odt, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if process.returncode != 0:
            print(f"LibreOffice ODT conversion error: {process.stderr}")
            return False
            
        # Now convert ODT to PDF with specific settings
        odt_file = input_file_abs.rsplit('.', 1)[0] + '.odt'
        cmd_to_pdf = [
            libreoffice_path,
            '--headless',
            '--convert-to',
            'pdf:draw_pdf_Export:{"IsSkipEmptyPages":{"type":"boolean","value":"false"},"Quality":{"type":"long","value":"100"},"ReduceImageResolution":{"type":"boolean","value":"false"},"UseTaggedPDF":{"type":"boolean","value":"true"},"SelectPdfVersion":{"type":"long","value":"1"},"ExportBookmarks":{"type":"boolean","value":"false"}}',
            '--outdir', output_dir,
            odt_file
        ]
        
        process = subprocess.run(cmd_to_pdf, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
        # Clean up the intermediate ODT file
        if os.path.exists(odt_file):
            os.remove(odt_file)
        
        if process.returncode != 0:
            print(f"LibreOffice PDF conversion error: {process.stderr}")
            return False
            
        output_file = input_file_abs.rsplit('.', 1)[0] + '.pdf'
        if os.path.exists(output_file):
            print(f"Successfully created PDF at {output_file}")
            return True
        else:
            print(f"PDF file was not created at {output_file}")
            return False
            
    except Exception as e:
        print(f"Error converting {input_file}: {str(e)}")
        return False

# converted = False
# # convert all doc/docx files to pdf
# for subdir in subdirs:
#     if not os.path.isdir(os.path.join(root_path, subdir)):
#         print(f"Not a directory: {subdir}")
#         continue
#     for file in os.listdir(os.path.join(root_path, subdir)):
#         if file.endswith(('.doc', '.docx')):
#             converted = True
#             file_path = os.path.join(root_path, subdir, file)
#             print(f"Converting {file} to pdf")
#             if convert_to_pdf_libreoffice(file_path):
#                 print(f"Successfully converted {file} to pdf")
#             else:
#                 print(f"Failed to convert {file}")
#                 # Add a longer delay between files if there was an error
#                 subprocess.run(['sleep', '3'])
#     # if converted:
#     #     break

# change the name of the pdf to the subdir name
# for subdir in subdirs:
#     if not os.path.isdir(os.path.join(root_path, subdir)):
#         print(f"Not a directory: {subdir}")
#         continue
#     # if there're more than one pdf, log error and continue
#     pdf_files = [f for f in os.listdir(os.path.join(root_path, subdir)) if f.endswith('.pdf')]
#     if len(pdf_files) > 1:
#         print(f"More than one pdf in {subdir}")
#         continue
#     elif len(pdf_files) == 0:
#         # print(f"No pdf in {subdir}")
#         continue
#     else:
#         pass
#         # shutil.move(os.path.join(root_path, subdir, pdf_files[0]), os.path.join(root_path, subdir, subdir + '.pdf'))
#         # print(f"Renamed {pdf_files[0]} to {subdir + '.pdf'}")

# move all doc/docx files to the 附件 folder
for subdir in subdirs:
    if not os.path.isdir(os.path.join(root_path, subdir)):
        print(f"Not a directory: {subdir}")
        continue
    os.makedirs(os.path.join(root_path, subdir, '附件'), exist_ok=True)
    for file in os.listdir(os.path.join(root_path, subdir)):
        if file.endswith(('.doc', '.docx')):
            shutil.move(os.path.join(root_path, subdir, file), os.path.join(root_path, subdir, '附件', file))
            print(f"Moved {file} to {os.path.join(root_path, subdir, '附件', file)}")
