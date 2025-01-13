from flask import Flask, request, jsonify
import requests
import json
import os
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# Your SharePoint site details
site_name = "coursemaster"
site_hostname = "aoscaustralia.sharepoint.com"

# Azure AD app registration details
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")
tenant_id = os.getenv("TENANT_ID")

# Function to get the access token from Azure AD
def get_access_token():
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
        "client_id": client_id,
        "client_secret": client_secret,
    }
    token_response = requests.post(token_url, data=token_data)
    return token_response.json().get("access_token") if token_response.status_code == 200 else None

# Function to get the site ID from the SharePoint site
def get_site_id(access_token):
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:/sites/{site_name}"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(site_url, headers=headers)
    return response.json().get("id") if response.status_code == 200 else None

# Function to create a folder in the "User directory" folder of the SharePoint document library
def create_folder(access_token, site_id, folder_name):
    # Define the path to the "User directory" folder in the document library
    user_directory_path = "User directory"  # Folder where the new folder will be created

    # Construct the full path for the new folder inside "User directory"
    folder_path = f"/{user_directory_path}/{folder_name}"

    # Modify the URL to point to the User directory folder inside the document library
    graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_path}:/"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    # Data for creating the folder
    folder_data = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename",  # To handle name conflicts
    }

    # Send the request to create the folder
    response = requests.put(graph_url, headers=headers, data=json.dumps(folder_data))

    # Return True if folder creation was successful, otherwise False
    return response.status_code == 201

# Route for handling the signup and folder creation
@app.route('/article', methods=['POST'])
def signup():
    data = request.json
    name = data.get('username')  # Get the username from the request body
    if not name:
        return jsonify({"message": "Name is required"}), 400

    # Get the access token from Azure AD
    access_token = get_access_token()
    if not access_token:
        return jsonify({"message": "Failed to get access token"}), 500

    # Get the SharePoint site ID
    site_id = get_site_id(access_token)
    if not site_id:
        return jsonify({"message": "Failed to get site ID"}), 500

    # Create the folder inside the "User directory"
    if create_folder(access_token, site_id, name):
        return jsonify({"message": f"Folder created successfully for {name}"}), 201
    else:
        return jsonify({"message": "Failed to create folder"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5010)
