from collections import defaultdict
from datetime import datetime
import time
from PyPDF2 import PdfReader
from os import path
from glob import glob
from openpyxl import load_workbook
import pandas as pd
import os
import pdfplumber
from ocr import find_name_amount_for_lohn, find_name_amount_for_lohnsteuer, find_name_amount_for_meldung, split_pdf

# CONFIG
#listOfEmployees = fetch_employees_data()
currentmonth = datetime.now().strftime("%m%y")

base_folder = "/Users/anirudhchawla/Library/CloudStorage/GoogleDrive-anirudhchawla@ming-group.de"

targetFolderFinal = f"{base_folder}/Shared drives/Ming Group/0_EmployeePaySlips"

# Function to find PDF files containing "Lohn" in their filenames
def find_lohn_pdfs(folder_path):
    print(folder_path)
    pdf_files = []
    for suffix in ["Lohnab", "Lohnsteuer", "Meldung"]:
        pdf_files.extend(
            [pdf for pdf in glob(os.path.join(folder_path, f"*{suffix}*.pdf"))
             if os.path.basename(pdf).count('_') <= 3]
        )
    print("Files found:", pdf_files)
    return pdf_files if pdf_files else None

# Set the company names and corresponding folder paths
company_info = {
    ##0924
    #"WIVH_GmbxH_WIVH": f"{base_folder}/Shared drives/Ming Group/1_WIVH_Wolfstreet Investment Holding GmbH/2_Accounting_WIVH/5_Salaries/" + currentmonth +"_WIVH",   
    #"Gastro_NF": f"{base_folder}/My Drive/2_1_0_NF_Ming_EK_NF_Gastroberatung/2_Accounting_NF/5_Salaries/"+ currentmonth + "_NF",
    "MICGMBH_MIC": f"{base_folder}/Shared drives/Ming Group/2_13_MIC_Ming Investment Consulting/2_Accounting_MIC/5_Salaries/"+ currentmonth + "_MIC",
    #"WSMGmbH_WSM": f"{base_folder}/Shared drives/Ming Group/2_12_WSM_Wolfstreet Management/2_Accounting_WSM/5_Salaries/"+ currentmonth +"_WSM",
    #"MingGastroGmbH_MT": f"{base_folder}/Shared drives/Ming Group/2_4_MT_MingGastroGmbH_Bikini/2_Accounting_MT/5_Salaries/" + currentmonth +"_MT",    
    #"ChenWangHandelsGmbH_FE": f"{base_folder}/Shared drives/Ming Group/2_3_FE_Feast/2_Accounting_FE/5_Salaries_FE/" + currentmonth +"_FE",
    #"MingDynastieGmbH_M2": f"{base_folder}/Shared drives/Ming Group/2_1_2_M2_Ming Dynastie GmbH Europa Center/2_Accounting_M2/5_Salaries/" + currentmonth + "_M2" ,
    #"CoffeeHanjanGmbH_HJ": f"{base_folder}/Shared drives/Ming Group/2_7_HJ_Coffee HanJan GmbH/2_Accounting_HJ/5_Salaries_HJ/" + currentmonth +"_HJ",
    #"KTVBarWolfgangFu_KTV": f"{base_folder}/Shared drives/Ming Group/2_5_KTV_EK_WOLF/2_Accounting_KTV/5_Salaries/" + currentmonth +"_KTV",
    #"HanFactoryGmbH_HF": f"{base_folder}/Shared drives/Ming Group/2_8_HF_Han Factory GmbH/2_Accounting_HF/5_Salaries_HF/"+ currentmonth +"_HF",
    #"BBMIGmbH_SSC": f"{base_folder}/Shared drives/Ming Group/2_9_1_SSC_BB Ming I GmbH SSC/2_Accounting_BBMI/5_Salaries/" + currentmonth +"_SSC",
    #"BBMIGmbH_AR": f"{base_folder}/Shared drives/Ming Group/2_9_2_AR_BB Ming I GmbH AROMA/2_Accounting_AR/5_Salaries/" + currentmonth +"_AR",
    #"MingJannoGmbH_M1": f"{base_folder}/Shared drives/Ming Group/2_1_1_M1_MingDynastieJannowitzbrueckeGmbH/2_Accounting_M1/5_Salaries/"+ currentmonth +"_M1",
    #"HANBBQ_WolfgangFu_H1": f"{base_folder}/Shared drives/Ming Group/2_6_H1_HANBBQ_EK/2_Accounting_H1/5_Salaries/" + currentmonth +"_H1"
}

import pdfplumber
from collections import defaultdict
import re

def extract_left_right_text(selected_pdf,suffix,target_folder,company):
    extracted_lines = []
    employee_list = []  # List to store employee dictionaries (name and page number)
    number_pattern = re.compile(r'\d')  # Pattern to identify lines with numbers
    input_pdf = PdfReader(open(selected_pdf, "rb"))  # Open PDF once
    with pdfplumber.open(selected_pdf) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):  # Start page numbering at 1
            extracted_lines.append(f"Page-{page_number}")  # Add the page number

            # Extract all words with their bounding boxes (coordinates)
            words = page.extract_words()

            # Separate words into left and right sides based on x-coordinate
            left_side = defaultdict(list)
            right_side = defaultdict(list)
            
            # Define the mid-point (assuming the page is divided in half by width)
            page_width = page.width
            mid_point = page_width / 2

            for word in words:
                # word['x0'] gives the x-coordinate of the text's left boundary
                y_coord = round(word['top'], 1)  # Use rounded y-coordinate to group words on the same line

                if word['x0'] < mid_point:  # Left side
                    left_side[y_coord].append(word['text'])
                else:  # Right side
                    right_side[y_coord].append(word['text'])

            # Sort the y-coordinates (lines) for both sides
            sorted_left_y_coords = sorted(left_side.keys())
            sorted_right_y_coords = sorted(right_side.keys())

            # Combine and store left and right side text
            page_lines = []  # List to hold the combined lines on this page

            # Append left side text first, combining words on the same line
            for y in sorted_left_y_coords:
                line_text = " ".join(left_side[y])
                page_lines.append(line_text)
                extracted_lines.append(line_text)

            # Append right side text next, combining words on the same line
            for y in sorted_right_y_coords:
                line_text = " ".join(right_side[y])
                page_lines.append(line_text)
                extracted_lines.append(line_text)

            # Now we apply the "Pers.-Nr" logic to find employee names on this page
            for i, line in enumerate(page_lines):
                # Search for "Pers.-Nr"
                if "Pers.-Nr" in line:
                    print(f"Found 'Pers.-Nr' on Page-{page_number}")
                    
                    # Start collecting lines until we find a line that does not have numbers (which will be the employee's name)
                    for j in range(i + 1, len(page_lines)):
                        if number_pattern.search(page_lines[j]):
                            # This line has numbers, skip it and continue
                            continue
                        else:
                            # The first line that does not have numbers is assumed to be the employee's name
                            name = page_lines[j].strip()
                            print(f"Found Employee Name: {name}")
                            split_pdf(currentmonth,input_pdf,page_number, target_folder, company, str(name).replace(" ",""), suffix)
                            # Add the employee name and page number to the list
                            employee_list.append({"Employee": name, "Page": page_number})
                            break  # Stop after finding the first name on this page
                    break
    return extracted_lines, employee_list

# Iterate over the company info and process each company
for company, folder_path in company_info.items():
    pdf_files = find_lohn_pdfs(folder_path)
    if pdf_files:
        target_folder = folder_path
        for selected_pdf in pdf_files:
            salary_data = []
            print(selected_pdf)
            # Extract text from the PDF once
            if "lohnab" in selected_pdf.lower() or "meldung" in selected_pdf.lower():
                with pdfplumber.open(selected_pdf) as pdf:
                    extracted_lines_for_amount = []
                    for page_number, page in enumerate(pdf.pages, start=1):  # Start page numbering at 1
                        text = page.extract_text()
                        if text:
                            lines = text.split('\n')
                            extracted_lines_for_amount.append(f"Page-{page_number}")
                            for line in lines:
                                extracted_lines_for_amount.append(line)  # Then append the line text
            start_time = time.time()
            if "lohnab" in selected_pdf.lower():
                suffix = "Lohnabrechnung"
                extracted_lines,employee_list = extract_left_right_text(selected_pdf,suffix,target_folder,company)
                for employee in employee_list:
                    target_word = 'Auszahlungsbetrag'  # Replace with the word you're looking for
                    find_name_amount_for_lohn(employee,extracted_lines_for_amount,target_word,company,salary_data,currentmonth)
                salary_dataframe = pd.DataFrame(salary_data)
                excel_file_name = f"{currentmonth}_{company}_NetSalaries.xlsx"
                file_path = path.join(target_folder, excel_file_name)
                # Check if the file exists
                if path.exists(file_path):
                    # Load existing data
                    existing_data = pd.read_excel(file_path)
                    
                    # Concatenate the existing data with the new data
                    updated_data = pd.concat([existing_data, salary_dataframe], ignore_index=True)
                else:
                    # If the file doesn't exist, use the new data as the updated data
                    updated_data = salary_dataframe

                # Write the updated data back to the file
                updated_data.to_excel(file_path, index=False)
            if "meldung" in selected_pdf.lower():
                suffix = "Meldungen"
                extracted_lines,employee_list = extract_left_right_text(selected_pdf,suffix,target_folder,company)
                print("employee_list")
                print(employee_list)
                for employee in employee_list:
                    target_word = "Bruttoarbeitsentgelt"
                    find_name_amount_for_meldung(employee,extracted_lines_for_amount,target_word,company,salary_data,currentmonth)
            if "lohnsteuer" in selected_pdf.lower():
                    suffix = "Lohnsteuerbescheinigung"
                    extracted_lines,employee_list = extract_left_right_text(selected_pdf,suffix,target_folder,company)
                #for employee in listOfEmployees:
                    """target_word = "Bruttoarbeitsentgelt"
                    find_name_amount_for_lohnsteuer(target_word,company,salary_data,currentmonth,extracted_lines,target_folder,suffix,input_pdf)"""
            end_time = time.time()
            print(f"Time taken for processing the employees: {end_time - start_time:.4f} seconds")  # Print the time taken
                #extract_text_from_pdf(selected_pdf, employee,company,salary_data,currentmonth)
            print("salary_data")
            print(salary_data)
            print(len(salary_data))
    else:
        print(f"No PDF files found for {company}.")

