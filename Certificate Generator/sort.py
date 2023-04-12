import os
import shutil
import pandas as pd

def organize_certificates(pdf_folder_path, excel_file_path):

    df = pd.read_excel(excel_file_path)

    # RUNS THROUGH EVERY PDF IN THE FOLDER
    for pdf_file_name in os.listdir(pdf_folder_path):

        # CHECKS IF THE FILE IS A PDF
        if pdf_file_name.endswith('.pdf'):

            # GETS THE PARTIAL NAME BETWEEN - AND .PDF
            partial_name = pdf_file_name.split(' - ')[-1].split('.')[0].title()

            # SEARCH FOR THE PARTIAL NAME IN THE SHEET
            matching_rows = df[df['NAME'] == partial_name]

            if matching_rows.empty:
                continue

            # CHECK IF THERE IS MORE THAN ONE MATCHING ROW
            if len(matching_rows) > 1:
                continue

            # GET THE ID AND THE EMAIL BY THE NAME

            folder_name = matching_rows.iloc[0]['ID']
            try:
                folder_name = 'ID_' + str(int(folder_name))
            except ValueError:
                continue

            # CREATE THE FOLDER IF IT DOES NOT EXIST
            folder_path = os.path.join(pdf_folder_path, folder_name)
            if not os.path.exists(folder_path):
                os.mkdir(folder_path)

            # MOVE THE PDF TO THE RIGHT FOLDER
            pdf_file_path = os.path.join(pdf_folder_path, pdf_file_name)
            new_pdf_file_path = os.path.join(folder_path, pdf_file_name)
            shutil.move(pdf_file_path, new_pdf_file_path)
