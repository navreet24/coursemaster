from flask import Flask, request, jsonify
from flask_cors import CORS
import msal
import requests
import os

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Configuration
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")
tenant_id = os.getenv("TENANT_ID")
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
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(endpoint, headers=headers)
    
    if response.status_code == 200:
        sites = response.json().get('value', [])
        for site in sites:
            if site_name.lower() in site.get('displayName', '').lower():
                return site.get('id')
        return None  # Site not found
    else:
        print(f"Error fetching site ID: {response.status_code} - {response.text}")
        return None

# Function to fetch document library contents
def get_document_library_contents(access_token, site_id, document_library_name):
    drives_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(drives_endpoint, headers=headers)
    
    if response.status_code == 200:
        drives = response.json().get('value', [])
        for drive in drives:
            if document_library_name.lower() in drive.get('name', '').lower():
                drive_id = drive.get('id')
                
                # Fetch the root folder contents
                folder_endpoint = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
                folder_response = requests.get(folder_endpoint, headers=headers)
                if folder_response.status_code == 200:
                    folders = folder_response.json().get('value', [])
                    
                    # Find "User directory" folder
                    for folder in folders:
                        if "User directory" in folder.get('name', ''):
                            return folder  # Or handle the folder as needed
                    
                    return {"error": "Folder 'User directory' not found"}  # Return error if not found
                    
                else:
                    print(f"Error fetching folder contents: {folder_response.text}")
                    return None
        print(f"Document library '{document_library_name}' not found")
        return None
    else:
        print(f"Error fetching drives: {response.text}")
        return None

# Function to fetch folder contents
def get_folder_contents(access_token, drive_id, folder_path):
    folder_endpoint = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}:/children"
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(folder_endpoint, headers=headers)

    if response.status_code == 200:
        return response.json().get('value', [])
    else:
        print(f"Error fetching folder contents: {response.text}")
        return None

# Flask route to fetch document library data
@app.route('/get-library-contents', methods=['POST'])
def get_library_contents():
    data = request.json
    site_name = "coursemaster"
    document_library_name = "Documents"

    if not site_name:
        return jsonify({"error": "Site name is required"}), 400

    access_token = get_access_token()
    if not access_token:
        return jsonify({"error": "Failed to obtain access token"}), 500

    site_id = get_site_id(access_token, site_name)
    if not site_id:
        return jsonify({"error": f"Site with name '{site_name}' not found"}), 404

    library_contents = get_document_library_contents(access_token, site_id, document_library_name)
    if library_contents:
        return jsonify({"success": True, "data": library_contents}), 200
    else:
        return jsonify({"error": "Failed to fetch document library contents"}), 500

# Flask route to open a folder
@app.route('/open-folder', methods=['POST'])
def open_folder():
    data = request.json
    site_name = data.get('site_name')
    document_library_name = data.get('document_library_name', 'Documents')  # Default to "Documents"
    folder_path = "User directory"  # Fixed to "User directory"
    username = data.get('username')  # Fetch the username from the request

    if not site_name or not username:
        return jsonify({"error": "Site name and username are required"}), 400

    access_token = get_access_token()
    if not access_token:
        return jsonify({"error": "Failed to obtain access token"}), 500

    site_id = get_site_id(access_token, site_name)
    if not site_id:
        return jsonify({"error": f"Site with name '{site_name}' not found"}), 404

    # Get drive ID of the document library
    drives_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    headers = {'Authorization': f'Bearer {access_token}'}
    drives_response = requests.get(drives_endpoint, headers=headers)

    if drives_response.status_code == 200:
        drives = drives_response.json().get('value', [])
        for drive in drives:
            if document_library_name.lower() in drive.get('name', '').lower():
                drive_id = drive.get('id')
                
                # Fetch folder contents of the "User directory"
                folder_contents = get_folder_contents(access_token, drive_id, folder_path)
                if folder_contents:
                    # Find the folder with the username inside the "User directory"
                    for folder in folder_contents:
                        if username.lower() == folder.get('name', '').lower():  # Match folder name with username
                            # Build the path to the user's folder
                            user_folder_path = f"{folder_path}/{folder['name']}"
                            user_folder_contents = get_folder_contents(access_token, drive_id, user_folder_path)
                            if user_folder_contents:
                                return jsonify({"success": True, "data": user_folder_contents}), 200
                            else:
                                return jsonify({"error": f"Failed to fetch contents of the user folder '{username}'"}), 500
                    return jsonify({"error": f"Folder with username '{username}' not found"}), 404
        return jsonify({"error": f"Document library '{document_library_name}' not found"}), 404
    else:
        return jsonify({"error": f"Error fetching drives: {drives_response.text}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)  # Run the application on port 5000
