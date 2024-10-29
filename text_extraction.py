from datetime import datetime
import fitz  # PyMuPDF
import re  # Regular expression module

def is_valid_date(date_string):
    # Regular expression to match the format dd.mm.yyyy
    date_pattern = r'^\d{2}\.\d{2}\.\d{4}$'
    
    # Check if the string matches the date format
    if re.match(date_pattern, date_string):
        # Try to parse the date to ensure it's a valid date
        try:
            datetime.strptime(date_string, "%d.%m.%Y")
            return True
        except ValueError:
            return False  # It's not a valid date if parsing fails
    return False  # It's not in the expected format

def extract_text_from_pdf(file_path, search_text,company,salary_data,currentmonth):
    with fitz.open(file_path) as pdf_doc:
        for page_num in range(pdf_doc.page_count):
            page = pdf_doc[page_num]
            text = page.get_text()
            
            # Remove spaces from each line
            lines = text.splitlines()  # Split text into lines
            processed_text = []
            for line in lines:
                line_no_spaces = line.replace(" ", "")  # Remove spaces from the line
                processed_text.append(line_no_spaces)

            # Check if the search text is in the processed text
            combined_text = "\n".join(processed_text)  # Combine processed lines
            if search_text in combined_text:
                # Find the last non-empty line of processed text
                last_non_empty_line = next((line for line in reversed(processed_text) if line.strip()), None)
                
                if last_non_empty_line:
                    print("last_non_empty_line")
                    print(f"Page {page_num + 1}: Last text found is: '{last_non_empty_line}'")
                    if len(last_non_empty_line) > 12 or is_valid_date(last_non_empty_line):
                        last_non_empty_line = ""
                    salary_data.append({"Employee": search_text, "NetSalary": last_non_empty_line, 
                                        "Verwendungszweck": f"{company} {currentmonth} Lohn/Gehalt" })
                    print(salary_data)
                else:
                    print(f"Page {page_num + 1}: No valid text found after '{search_text}'.")

pdf_file_path = "/Users/anirudhchawla/Library/CloudStorage/GoogleDrive-anirudhchawla@ming-group.de/Shared drives/Ming Group/2_9_2_AR_BB Ming I GmbH AROMA/2_Accounting_AR/5_Salaries/0924_AR/Lohnabrechnungen AR_09_2024.pdf"
employeename = "HuyGiapBui"  # Replace with the specific text you want to search for

