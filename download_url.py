import base64
import os
import webbrowser

import msal
import pyperclip
import requests


CLIENT_ID = "ab9b8c07-8f02-4f72-87fa-80105867a763"  # OneDrive sync client
SCOPES = ["https://graph.microsoft.com/Files.Read"]


def encode_sharing_url(url):
    # Microsoft Graph API requires base64url encoding with u! prefix
    encoded = base64.urlsafe_b64encode(url.encode()).rstrip(b"=").decode()
    return "u!" + encoded


def get_access_token():
    app = msal.PublicClientApplication(
        CLIENT_ID, authority="https://login.microsoftonline.com/common"
    )

    # try silent refresh first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    # fall back to device code flow
    flow = app.initiate_device_flow(scopes=SCOPES)
    webbrowser.open("https://microsoft.com/devicelogin")
    print(f"Microsoft: {flow['message']}")
    pyperclip.copy(flow["user_code"])
    print(f"Code: {flow['user_code']} copied to clipboad.")

    # return access token
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", "auth failed"))
    return result["access_token"]


def get_root_item(sharing_url, token):
    encoded = encode_sharing_url(sharing_url)
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{encoded}/driveItem",
        headers=headers,
        params={"$select": "id,name,parentReference"},
    )
    resp.raise_for_status()
    return resp.json()


def list_children(item_id, drive_id, token):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"
    params = {"$select": "id,name,file,folder"}
    headers = {"Authorization": f"Bearer {token}"}

    while url:
        resp = requests.get(url, headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
        yield from data.get("value", [])
        url = data.get("@odata.nextLink")
        params = None


def download_folder(sharing_url, token, dest_dir):
    root = get_root_item(sharing_url, token)
    drive_id = root["parentReference"]["driveId"]
    _download_children(root["id"], drive_id, token, dest_dir)


def _download_children(item_id, drive_id, token, dest_dir):
    os.makedirs(dest_dir, exist_ok=True)
    headers = {"Authorization": f"Bearer {token}"}

    for item in list_children(item_id, drive_id, token):
        dest_path = os.path.join(dest_dir, item["name"])

        if "folder" in item:
            _download_children(item["id"], drive_id, token, dest_path)
        elif "file" in item:
            content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item['id']}/content"
            with requests.get(content_url, headers=headers, stream=True) as r:
                r.raise_for_status()
                with open(dest_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)


with open("anyone.url", "r") as url_file:
    url_str = url_file.readline()

token = get_access_token()
download_folder(url_str, token, os.path.abspath("./down"))
