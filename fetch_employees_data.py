import requests

def fetch_employees_data():
    # URL of the deployed Google Apps Script Web App
    web_app_url = "https://script.google.com/macros/s/AKfycbwI_LLBt2i3C430o_Cqn1TBK_kOw6dAFacuoYNBqmRXuHV1bfVpb4daT1xsLqvB4Li3/exec"  # Replace with your web app URL

    try:
        # Send a GET request to the Google Apps Script
        response = requests.get(web_app_url)
        
        if response.status_code == 200:
            # Parse the JSON response
            data = response.json()
            print("Fetched Data:", data)
            return data
        else:
            print(f"Failed to fetch data. Status code: {response.status_code}")
    except Exception as e:
        print(f"An error occurred: {e}")

