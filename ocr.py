from pathlib import Path
import re
import time
from PyPDF2 import PdfWriter
from os import path
import os as os

cloud_storage_folder = Path(os.path.expanduser("~")) / "Library" / "CloudStorage"

# Find the Google Drive folder that matches the "GoogleDrive-<email>@ming-group.de" pattern
base_folder = None
for folder in cloud_storage_folder.iterdir():
    if folder.is_dir() and folder.name.startswith("GoogleDrive-") and folder.name.endswith("@ming-group.de"):
        base_folder = folder
        break
targetFolderFinal = f"{base_folder}/Shared drives/Ming Group/0_EmployeePaySlips"

def find_name_amount_for_lohn(employeedetails, extracted_lines_for_amount, target_word, company, salary_data, currentmonth):
    lines = extracted_lines_for_amount
    name = employeedetails['Employee']
    page_number_to_process = employeedetails['Page']
    
    # Regex to match lines that contain only numbers
    number_pattern = re.compile(r'\d')

    for i, line in enumerate(lines):
        if "Page" in line:
            # Extract the page number from the line
            page_number = int(line.split("-")[1])  # Assuming the format is "Page X"
            
            # Process only the page corresponding to the employee's page number
            if page_number == page_number_to_process:
                print(f"Processing Page-{page_number} for Employee: {name}")
                
                # Now look for the target word (e.g., salary-related info)
                next_page_index = len(lines)  # Default to the end of the document if no more pages
                target_word_count = 0  # Counter to track the number of times target word is found
                gross_salary_found = False  # Flag to track if gross salary has been found

                for k in range(i + 1, len(lines)):
                    if "Page" in lines[k]:
                        next_page_index = k  # The next page starts here
                        break
                # Initialize variables to hold net and gross salary
                net_salary = None
                gross_salary = None

                # Search for the target word and "Gesamt-Brutto" within this page's lines
                for l in range(i + 1, next_page_index):
                    if target_word in lines[l]:
                        target_word_count += 1  # Increment the counter when the target word is found
                        
                        # Check the line below the target word
                        if l + 1 < len(lines):
                            line_below = lines[l + 1]
                            words = line_below.split()
                            print(f"Words found after target word: {words}")
                            if words:
                                net_salary = words[-1] # Remove trailing '-' if present
                                if target_word_count == 1:
                                    # First occurrence: prepare to add salary data
                                    print(f"Net salary found for {name} on first occurrence: {net_salary}")
                                elif target_word_count == 2:
                                    # Second occurrence: we can overwrite with the new net salary
                                    print(f"Net salary found for {name} on second occurrence: {net_salary}")

                    # Logic for "Gesamt-Brutto"
                    if "Gesamt-Brutto" in lines[l] and not gross_salary_found:
                        gross_salary_found = True  # Set the flag to indicate gross salary is found
                        
                        # Check the line below "Gesamt-Brutto"
                        if l + 1 < len(lines):
                            line_below = lines[l + 1]
                            words = line_below.split()
                            print(f"Words found after 'Gesamt-Brutto': {words}")
                            if words:
                                gross_salary = words[-1]  # Remove trailing '-' if present
                                print(f"Gross salary found for {name}: {gross_salary}")
                                

                # After processing, add to salary_data if we found both net and gross salary
                if net_salary is not None or gross_salary is not None:
                    salary_entry = {
                        "Employee": name,
                        "NetSalary": net_salary if net_salary is not None else None,
                        "GrossSalary": gross_salary if gross_salary is not None else None,
                    }
                    
                    # Add 'Verwendungszweck' as the last key
                    salary_entry["Verwendungszweck"] = f"{company} {currentmonth} Lohn/Gehalt"
                    
                    salary_data.append(salary_entry)
                    print(f"Salary data added for {name}: {salary_entry}")

    return "Processing complete."

def find_name_amount_for_meldung(employeedetails, extracted_lines, target_word, company, salary_data, currentmonth):
    lines = extracted_lines
    name = employeedetails['Employee']  # Use the employee name from employeedetails
    page_number_to_process = employeedetails['Page']  # Use the page number from employeedetails
    euro_pattern = re.compile(r"Euro")  # Regex to match lines that contain "Euro"

    for i, line in enumerate(lines):
        if "Page" in line:
            # Extract the page number from the line
            page_number = int(line.split("-")[1])  # Assuming the format is "Page X"
            
            # Process only the page corresponding to the employee's page number
            if page_number == page_number_to_process:
                print(f"Processing Page-{page_number} for Employee: {name}")
                
                # Now look for the target word (e.g., salary-related info)
                next_page_index = len(lines)  # Default to the end of the document if no more pages

                for k in range(i + 1, len(lines)):
                    if "Page" in lines[k]:
                        next_page_index = k  # The next page starts here
                        break
                
                # Search for the target word within this page's lines
                for j in range(i + 1, next_page_index):
                    if target_word in lines[j]:
                        # Check if "Euro" is in the same line
                        if euro_pattern.search(lines[j]):
                            words = lines[j].split()
                            if len(words) > 1:
                                amount = words[1]  # Assuming the format "target_word Euro <amount>"
                                if amount != "Euro":
                                    salary_data.append({
                                        "Employee": name,
                                        "NetSalary": amount,
                                        "Verwendungszweck": f"{company} {currentmonth} Lohn/Gehalt"
                                    })
                                    print(f"Salary data added for {name} with amount: {amount}")
                                else:
                                    salary_data.append({
                                        "Employee": name,
                                        "NetSalary": "",
                                        "Verwendungszweck": f"{company} {currentmonth} Lohn/Gehalt"
                                    })
                                    print(f"Salary data added for {name} with empty amount")
                            else:
                                print(f"No salary amount found for {name} in the line containing Euro.")
                        else:
                            print(f"Euro not found in the line containing {target_word} for {name}.")

    print("Processing complete.")


def find_name_amount_for_lohnsteuer(target_word, company, salary_data, currentmonth, extracted_lines, target_folder, suffix, input_pdf):
    lines = extracted_lines
    # Regex to match lines that contain only numbers
    number_pattern = re.compile(r'\d')

    for i, line in enumerate(lines):
        if "Page" in line:
            # Extract the page number from the line
            page_number = int(line.split("-")[1])  # Assuming the format is "Page X"
            print(f"Processing Page-{page_number}")
        
        # Search for "Hinweise zur Abrechnung"
        if "Pers.-Nr" in line:
            print(f"Found 'Pers.-Nr' on Page-{page_number}")
            
            # Start collecting lines until we find a line that does not have numbers (which will be the employee's name)
            for j in range(i + 1, len(lines)):
                    # The first line that does not have numbers is assumed to be the employee's name
                    name = lines[i + 1].strip()
                    print(f"Found Employee Name: {name}")
                    # Now look for the target word (e.g., salary-related info) after this
                    next_page_index = len(lines)  # Default to the end of the document if no more pages
                    for k in range(j + 1, len(lines)):
                        if "Page" in lines[k]:
                            next_page_index = k  # The next page starts here
                            break
                    #split_pdf(input_pdf,page_number, target_folder, company, name, suffix)
                    break  # Stop processing after finding the name and salary info

    return "Processing complete."

def split_pdf(currentmonth,input_pdf, page_number,target_folder, company, employee, suffix):
    start_time = time.time()
    
    # Construct the filename
    filename = f"{company}_{currentmonth}_{employee}_{page_number-1}_{suffix}"
    file_path = path.join(target_folder, f"{filename}.pdf")
    print("filename")
    print(filename)
    print(file_path)
    # Check if the PDF file already exists
    if path.exists(file_path):
        print(f"{file_path} already exists. Skipping write operation.")
        elapsed_time = time.time() - start_time  # End time
        print(f"split_pdf took {elapsed_time:.2f} seconds")
        return  # Skip if the file already exists

    output = PdfWriter()
    print(page_number)
    page = input_pdf.pages[page_number-1]

    # Check if the page is empty
    if not page.extract_text().strip():  # Skip if the page is empty
        print(f"Page {page_number} is empty. Skipping.")
        elapsed_time = time.time() - start_time  # End time
        print(f"split_pdf took {elapsed_time:.2f} seconds")
        return
    
    output.add_page(page)
    print("emaployee name")
    print(employee)
    # Write the PDF to the file
    with open(file_path, "wb") as output_stream:
        print(f"Writing to {file_path}")
        output.write(output_stream)

    # Create employee directory if not exists (do this only once per employee)
    employee_folder = path.join(targetFolderFinal, employee)
    #Path(employee_folder).mkdir(parents=True, exist_ok=True)

    # Save a copy in the employee's folder, only if not already existing
    employee_file_path = path.join(employee_folder, f"{filename}.pdf")
    if not path.exists(employee_file_path):
        with open(employee_file_path, "wb") as output_stream:
            print(f"Writing to {employee_file_path}")
            output.write(output_stream)

    elapsed_time = time.time() - start_time  # End time
    print(f"split_pdf took {elapsed_time:.2f} seconds")
