from flask import Flask, request, jsonify
from flask_cors import CORS
import msal
import requests

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Configuration
client_id = 'b25fac3a-8c66-46af-b161-840828c538f3'
client_secret = 'GFm8Q~q6lG0UO3_ZT-~LVMNjgvEND6~cRXt0pc_f'
tenant_id = '4dab0fef-f02d-440b-97c3-712e9483bd68'
authority = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ["https://graph.microsoft.com/.default"]

# MSAL Application
msal_app = msal.ConfidentialClientApplication(
    client_id=client_id,
    client_credential=client_secret,
    authority=authority
)

# Helper function to get access token
def get_access_token():
    result = msal_app.acquire_token_silent(scopes, account=None)
    if not result:
        result = msal_app.acquire_token_for_client(scopes=scopes)
    if "access_token" in result:
        return result['access_token']
    else:
        return None

# Function to fetch Site ID
def get_site_id(access_token, site_name):
    endpoint = 'https://graph.microsoft.com/v1.0/sites/'
    response = requests.get(endpoint, headers={'Authorization': f'Bearer {access_token}'})
    
    if response.status_code == 200:
        sites = response.json().get('value', [])
        for site in sites:
            if site_name.lower() in site.get('displayName', '').lower():
                return site.get('id')
        return None  # Site not found
    else:
        return None  # API call failed

# Function to fetch List ID
def get_list_id(access_token, site_id, list_name):
    endpoint = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists'
    response = requests.get(endpoint, headers={'Authorization': f'Bearer {access_token}'})
    
    if response.status_code == 200:
        lists = response.json().get('value', [])
        for lst in lists:
            if list_name.lower() in lst.get('displayName', '').lower():
                return lst.get('id')
        return None  # List not found
    else:
        return None  # API call failed

# Function to fetch List Items
def get_list_items(access_token, site_id, list_id):
    endpoint = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?expand=fields'
    response = requests.get(endpoint, headers={'Authorization': f'Bearer {access_token}'})
    
    if response.status_code == 200:
        return response.json().get('value', [])  # Return list items
    else:
        return None  # API call failed

# Function to extract titles from list items and include ID
def extract_titles_from_items(list_items):
    # Extract the 'Title' and 'ID' fields from each item's fields
    return [{"id": item['id'], "title": item['fields']['Title']} for item in list_items if 'Title' in item['fields']]

# Function to fetch Units based on Qualification ID
def get_units_by_qualification(access_token, site_id, qualification_id):
    list_name = "Units"  # Assuming your unit list is called "Units"
    list_id = get_list_id(access_token, site_id, list_name)
    if not list_id:
        return None  # List not found

    # Fetch Units for the specific qualification
    endpoint = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?$filter=fields/QualificationId eq {qualification_id}&expand=fields'
    response = requests.get(endpoint, headers={'Authorization': f'Bearer {access_token}'})
    
    if response.status_code == 200:
        return response.json().get('value', [])
    else:
        return None  # API call failed

# Main function to handle the API logic
def fetch_site_list_data_and_items(site_name, list_name):
    # Step 1: Get access token
    access_token = get_access_token()
    if not access_token:
        print("Failed to obtain access token")
        return {"error": "Failed to obtain access token"}, 500
    
    # Step 2: Fetch Site ID
    site_id = get_site_id(access_token, site_name)
    if not site_id:
        print(f"Site with name '{site_name}' not found")
        return {"error": f"Site with name '{site_name}' not found"}, 404
    
    # Step 3: Fetch List ID
    list_id = get_list_id(access_token, site_id, list_name)
    if not list_id:
        print(f"List with name '{list_name}' not found on site '{site_name}'")
        return {"error": f"List with name '{list_name}' not found on site '{site_name}'"}, 404
    
    # Step 4: Fetch List Items
    list_items = get_list_items(access_token, site_id, list_id)
    if list_items is None:
        print("Failed to fetch list items")
        return {"error": "Failed to fetch list items"}, 500
    
    # Step 5: Extract Titles from the List Items
    titles = extract_titles_from_items(list_items)
    
    # Return results
    return {
       "site_name": site_name,
        "list_name": list_name,
        "site_id": site_id,
        "list_id": list_id,
        "titles": titles,  # Include both ID and title
        "list_items": list_items  # Optional: Include raw list items
    }, 200


# Flask route for POST request to fetch qualifications
@app.route('/get-qualificationid', methods=['POST'])
def get_qualification():
    data = request.json
    site_name = data.get('site_name')
    list_name = data.get('list_name')
    
    if not site_name or not list_name:
        return jsonify({"error": "Both 'site_name' and 'list_name' are required"}), 400
    
    result, status_code = fetch_site_list_data_and_items(site_name, list_name)
    return jsonify(result), status_code


# Flask route for POST request to fetch units by qualification ID


@app.route('/get-units', methods=['POST'])
def get_units():
    data = request.json
    qualification_id = data.get('qualification_id')
    site_name = data.get('site_name')

    if not qualification_id or not site_name:
        return jsonify({"error": "Both 'qualification_id' and 'site_name' are required"}), 400

    # Get access token
    access_token = get_access_token()
    if not access_token:
        return jsonify({"error": "Failed to obtain access token"}), 500

    # Get Site ID
    site_id = get_site_id(access_token, site_name)
    if not site_id:
        return jsonify({"error": f"Site with name '{site_name}' not found"}), 404

    # Fetch units based on qualification ID
    units = get_units_by_qualification(access_token, site_id, qualification_id)
    if not units:
        return jsonify({"error": "Failed to fetch units"}), 500

    return jsonify({"units": units}), 200



@app.route('/get-learner-guides', methods=['POST'])
def fetch_learner_guides():
    data = request.json
    site_name = data.get('site_name')

    if not site_name:
        return jsonify({"error": "'site_name' is required"}), 400

    # Get access token
    access_token = get_access_token()
    if not access_token:
        return jsonify({"error": "Failed to obtain access token"}), 500

    # Get Site ID
    site_id = get_site_id(access_token, site_name)
    if not site_id:
        return jsonify({"error": f"Site with name '{site_name}' not found"}), 404

    # Fetch Learner Guides
    learner_guides = fetch_learner_guides(access_token, site_id)
    if not learner_guides:
        return jsonify({"error": "Failed to fetch learnerguide"}), 500

    return jsonify({"learner_guides": learner_guides}), 200

if __name__ == '__main__':
    app.run(debug=True, port=5002)  # Port is now 5002
