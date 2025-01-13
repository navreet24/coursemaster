import requests
from flask import Flask, request, jsonify
from io import BytesIO
import PyPDF2
from docx import Document
from datetime import datetime
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from flask_cors import CORS


# Initialize the Flask app
app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "http://localhost:5174"}})

# API Configuration
openrouter_api_key = "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"  # Replace with your API key
client_id = 'b25fac3a-8c66-46af-b161-840828c538f3'
client_secret = 'GFm8Q~q6lG0UO3_ZT-~LVMNjgvEND6~cRXt0pc_f'
tenant_id = '4dab0fef-f02d-440b-97c3-712e9483bd68'

# Helper Functions
def get_access_token():
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    body = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(token_url, headers=headers, data=body)
    if response.status_code == 200:
        return response.json().get('access_token')
    else:
        raise Exception(f"Failed to obtain access token: {response.text}")

def get_site_id(site_url, access_token):
    full_url = f'https://graph.microsoft.com/v1.0/sites/{site_url}'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(full_url, headers=headers)
    response.raise_for_status()
    return response.json().get('id')

def get_drive_id(site_id, access_token):
    drives_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(drives_url, headers=headers)
    response.raise_for_status()
    drives = response.json().get('value', [])
    if not drives:
        raise ValueError("No drives found for the given site.")
    return drives[0]['id']

def list_folder_contents(site_id, drive_id, access_token, folder_id='root'):
    folder_contents_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_id}/children'
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(folder_contents_url, headers=headers)
    response.raise_for_status()
    folder_contents = response.json()
    items_list = []
    if 'value' in folder_contents:
        for item in folder_contents['value']:
            if 'folder' in item:
                items_list.append((item['id'], item['name'], 'Folder'))
            elif 'file' in item:
                items_list.append((item['id'], item['name'], 'File'))
    return items_list

def download_and_extract_pdf_text(site_id, drive_id, file_id, access_token):
    download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}/content"
    headers = {'Authorization': f'Bearer {access_token}'}
    try:
        response = requests.get(download_url, headers=headers)
        response.raise_for_status()
        pdf_file = BytesIO(response.content)
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        extracted_text = ""
        for page in pdf_reader.pages:
            extracted_text += page.extract_text()
        return extracted_text
    except Exception as e:
        print(f"Error processing PDF: {e}")
        return None

def chunk_text(text, chunk_size=3000):
    return [text[i:i + chunk_size] for i in range(0, len(text), chunk_size)]

def create_word_document(content,unit_name):
    filename = f"{unit_name}_summative.docx"
    
    # Create the Word document
    doc = Document()
    doc.add_heading(" Assessment", level=1)
    for section in content.split("\n\n"):
        doc.add_paragraph(section)
    doc.save(filename)
    return filename

def find_folder_id_from_path(site_id, drive_id, folder_path, access_token):
    parts = folder_path.strip('/').split('/')
    folder_id = 'root'
    for part in parts:
        contents = list_folder_contents(site_id, drive_id, access_token, folder_id)
        next_folder_id = None
        for item_id, item_name, item_type in contents:
            if item_type == 'Folder' and item_name == part:
                next_folder_id = item_id
                break
        if next_folder_id is None:
            raise ValueError(f"Folder '{part}' not found in the path.")
        folder_id = next_folder_id
    return folder_id

def extract_data_under_heading(soup, heading_text):
    headings = soup.find_all(lambda tag: tag.name in ["h1", "h2", "h3", "h4"] and heading_text.lower() in tag.get_text(strip=True).lower())
    if headings:
        heading = headings[0]
        container = heading.find_next('div', {'class': 'table-std'}) or heading.find_next('div')
        if container:
            rows = container.find_all('tr')
            if rows:
                data = []
                for row in rows:
                    cols = row.find_all('td')
                    if cols:
                        data.append([col.get_text(strip=True) for col in cols])
                return data
            else:
                return container.get_text(strip=True)
        else:
            return None
    return None

def scrape_training_details(unit_code):
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--incognito")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    url = f"https://training.gov.au/training/details/{unit_code}"
    driver.get(url)

    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.table-std'))
        )
    except Exception as e:
        print("Timeout error:", e)

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    elements_and_performance_criteria = extract_data_under_heading(soup, "Elements and Performance Criteria")
    knowledge_evidence = extract_data_under_heading(soup, "Knowledge Evidence")

    driver.quit()
    return knowledge_evidence, elements_and_performance_criteria

def upload_file_to_sharepoint(site_id, drive_id, folder_id, file_path, access_token):
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_id}:/{file_path.split('/')[-1]}:/content"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/octet-stream'
    }
    with open(file_path, 'rb') as file:
        response = requests.put(upload_url, headers=headers, data=file)
    response.raise_for_status()
    return response.json()

@app.route('/generate_assessment1', methods=['POST'])
def generate_assessment():
    try:
        data = request.get_json()
        print(f"Incoming request data: {data}")

        # Retrieve username and selected units
        Username = data.get('username')
        if not Username:
            return jsonify({'error': 'Username is required in the request payload'}), 400
        
        selected_units = data.get('selected_units', [])
        if not selected_units:
            return jsonify({'error': 'Selected units are required'}), 400
        
        unit_code = selected_units[0] if selected_units else None
        if not unit_code:
            return jsonify({'error': 'Unit code is required'}), 400

        # Get the unit name (assuming the unit_code itself is the name, adjust as necessary)
        unit_name = unit_code  # You can modify this to get the unit name if available
        site_url = "aoscaustralia.sharepoint.com:/sites/coursemaster"
        folder_path = "learner guides"
         #Use the first unit code or process as needed
        base_directory = "User directory"


        # if not site_url or not folder_path or not unit_code:
        #     return jsonify({'error': 'Site URL, folder path, and unit code are required'}), 400

        access_token = get_access_token()
        site_id = get_site_id(site_url, access_token)
        drive_id = get_drive_id(site_id, access_token)
        folder_id = find_folder_id_from_path(site_id, drive_id, folder_path, access_token)
        folder_contents = list_folder_contents(site_id, drive_id, access_token, folder_id)

        pdf_content = None
        for item in folder_contents:
            if item[2] == 'File' and item[1].endswith('.pdf'):
                pdf_content = download_and_extract_pdf_text(site_id, drive_id, item[0], access_token)
                break

        if not pdf_content:
            return jsonify({'error': 'No PDF files found in the folder'}), 404

        knowledge_evidence, performance_criteria = scrape_training_details(unit_code)
        chunks = chunk_text(pdf_content)
        assessments = []

        for idx, chunk in enumerate(chunks):
            combined_content = {
                "knowledge_evidence": knowledge_evidence,
                "performance_criteria": performance_criteria
            }

            messages = [
                {"role": "system", "content": "You are an educational assistant that generates assessments from provided learner guides, knowledge evidence, and performance criteria."},
                {"role": "user", "content": f"Using the following knowledge evidence and performance criteria, generate an assessment that includes class activities and questions like role-play, simulation, oral interview questions, group activity, demonstration:\n\nKnowledge Evidence:\n{chunk}\n\nPerformance Criteria:\n{chunk}"
            }
            ]

            while True:
                try:
                    response = requests.post(
                        "https://openrouter.ai/api/v1/chat/completions",
                        headers={
                            "Authorization": f"Bearer {openrouter_api_key}",
                            "Content-Type": "application/json",
                        },
                        json={
                            "model": "meta-llama/llama-3.2-3b-instruct:free",
                            "messages": messages,
                            "max_tokens": 10000,
                            "temperature": 0.7,
                        },
                    )

                    if response.status_code == 429:
                        retry_after = int(response.headers.get("X-RateLimit-Reset", time.time())) - int(time.time())
                        print(f"Rate limit exceeded. Retrying after {retry_after} seconds...")
                        time.sleep(retry_after)
                    else:
                        break
                except requests.exceptions.RequestException as e:
                    print(f"Error making request: {e}")
                    return jsonify({'error': f"Request error: {e}"}), 500

            try:
                response_data = response.json()

                if 'choices' in response_data:
                    assessment = response_data['choices'][0]['message']['content'].strip()
                    assessments.append(assessment)
                else:
                    assessments.append(f"Failed to generate assessment for Chunk {idx + 1}. No 'choices' in the response.")
            except (KeyError, IndexError) as e:
                print(f"Error parsing API response: {e}")
                assessments.append(f"Failed to parse API response for Chunk {idx + 1}.")

        combined_assessment = "\n\n".join(assessments)
        word_filename = create_word_document(combined_assessment,unit_name)

        try:
            user_directory_path = f"{base_directory}/{Username}"
            user_folder_id = find_folder_id_from_path(site_id, drive_id, user_directory_path, access_token)

            upload_file_to_sharepoint(site_id, drive_id, user_folder_id, word_filename, access_token)

            return jsonify({'assessment': combined_assessment, 'word_file': word_filename, 'upload_status': 'Success'}), 200
        except Exception as upload_error:
            return jsonify({'assessment': combined_assessment, 'word_file': word_filename, 'upload_status': 'Failed', 'error': str(upload_error)}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5019)