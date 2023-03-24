import os
import shutil
import pandas as pd
import tkinter as tk
import glob
import traceback
from tkinter import filedialog
import win32com.client as win32

# USER SELECTION
print('Select the folder with the certificates')
pdf_folder_path = filedialog.askdirectory()
print('Path: ', pdf_folder_path)

print('Select the excel file with the info: ')
excel_file_path = filedialog.askopenfilename()
print('Path: ', excel_file_path)

df = pd.read_excel(excel_file_path)
email_list = {}

# RUNS THROUGH EVERY PDF IN THE FOLDER
for pdf_file_name in os.listdir(pdf_folder_path):

    # CHECKS IF THE FILE IS A PDF
    if pdf_file_name.endswith('.pdf'):

        # GETS THE PARTIAL NAME BETWEEN - AND .PDF
        partial_name = pdf_file_name.split(' - ')[-1].split('.')[0]

        # SEARCH FOR THE PARTIAL NAME IN THE SHEET
        matching_rows = df[df['NAME'] == partial_name]

        if matching_rows.empty:
            continue

        # CHECK IF THERE IS MORE THAN ONE MATCHING ROW
        if len(matching_rows) > 1:
            print(f"Person '{partial_name}' has more than one entry.")
            continue

        # GET THE ID AND THE EMAIL BY THE NAME
        folder_name = matching_rows.iloc[0]['ID']
        email_address = matching_rows.iloc[0]['EMAIL']

        folder_name = 'ID_' + str(folder_name)
        email_address = str(email_address)

        email_list[folder_name] = email_address

        # CREATE THE FOLDER IF IT DOES NOT EXIST
        folder_path = os.path.join(pdf_folder_path, folder_name)
        if not os.path.exists(folder_path):
            os.mkdir(folder_path)

        # MOVE THE PDF TO THE RIGHT FOLDER
        pdf_file_path = os.path.join(pdf_folder_path, pdf_file_name)
        new_pdf_file_path = os.path.join(folder_path, pdf_file_name)
        shutil.move(pdf_file_path, new_pdf_file_path)

folders = [f.path for f in os.scandir(pdf_folder_path) if f.is_dir()]

# RUNS THROUGH EVERY FOLDER THAT STARTS WITH ID_ INSIDE THE USER SELECTED FOLDER
for folder in folders:

    folder_name = os.path.basename(folder)
    if not folder_name.startswith('ID_'):
        continue

    try:
        # STARTS OUTLOOK INSTANCE
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        root = tk.Tk()
        root.withdraw()

        # GETS THE EMAIL BY THE ID INSIDE THE EMAIL_ADRESS ARRAY
        email_address = df.loc[df['ID'] == folder_name, 'EMAIL'].iloc[0]

        if not email_address:
            continue

        # CREATE THE EMAIL
        mail.To = email_address
        
        # TRY TO GET THE EMAIL TEMPLATE IN THE TXT FILE, WRITE DEFAULT SUBJECT AND BODY IF NOT FOUND
        try:
            with open('email_template.txt', 'r', encoding='utf-8') as f:
                content = f.read()
                mail.subject = content.split('[Subject] = ')[1].split('\n')[0]
                mail.body = content.split('[Body] = ')[1]
        except FileNotFoundError:
            mail.subject = 'Certificate'
            mail.body = 'Hi.\n\n Here is your certificate.\n\nxoxo <3'

        # ATTACH ALL PDF FILES INSIDE THE CURRENT FOLDER
        pdf_files = glob.glob(os.path.join(folder, '*.pdf'))
        for pdf_file in pdf_files:
            mail.Attachments.Add(pdf_file)
        mail.Send()

        print('Email sent to: ', email_address)

    # IF SOMETHING GOES WRONG, IT WILL WRITE THE ERROR IN A LOG FILE
    except Exception as e:
        with open('error.log', 'a') as f:
            f.write(str(e))
            f.write(traceback.format_exc())