import base64
import datetime
import os
import time
import webbrowser

import msal
import pyperclip
import requests


client_id = "ab9b8c07-8f02-4f72-87fa-80105867a763"  # OneDrive sync client
scopes = ["https://graph.microsoft.com/Files.Read"]
dest_root = os.path.abspath("./down")


def _download_children(item_id, drive_id, app, session, dest_dir, dest_root):
    os.makedirs(dest_dir, exist_ok=True)
    print(f"FOLDER: {os.path.relpath(dest_dir, dest_root)}")
    token = refresh_token(app)
    for item in list_children(item_id, drive_id, token):
        dest_path = os.path.join(dest_dir, item["name"])
        if "folder" in item:
            _download_children(item["id"], drive_id, app, session, dest_path, dest_root)
        elif "file" in item:
            download_file(item, drive_id, dest_path, dest_root, app, session)


def download_file(item, drive_id, dest_path, dest_root, app, session):
    expected_size = item["size"]
    expected_mtime = parse_mtime(item["lastModifiedDateTime"])
    if is_complete(dest_path, expected_size, expected_mtime):
        print(f"(SKIP): {os.path.relpath(dest_path, dest_root)}")
        return
    # refresh token + downloadUrl per file (both expire ~1hr)
    token = refresh_token(app)
    fresh = get_item(item["id"], drive_id, token)
    print(f"FILE: {os.path.relpath(dest_path, dest_root)}  ({expected_size:,} bytes)")
    stream_with_retry(
        fresh["@microsoft.graph.downloadUrl"], dest_path, expected_mtime, session
    )


def download_folder(sharing_url, app, dest_dir):
    token = refresh_token(app)
    root = get_root_item(sharing_url, token)
    drive_id = root["parentReference"]["driveId"]
    dest_root = os.path.dirname(dest_dir)
    with requests.Session() as session:
        _download_children(root["id"], drive_id, app, session, dest_dir, dest_root)


def encode_sharing_url(url):
    # Microsoft Graph API requires base64url encoding with u! prefix
    encoded = base64.urlsafe_b64encode(url.encode()).rstrip(b"=").decode()
    return "u!" + encoded


def get_access_token():
    app = msal.PublicClientApplication(
        client_id, authority="https://login.microsoftonline.com/common"
    )
    # try silent refresh first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            return app
    # fall back to device code flow
    flow = app.initiate_device_flow(scopes=scopes)
    webbrowser.open("https://microsoft.com/devicelogin")
    print(f'"{flow["message"]}" - Microsoft', end=2 * os.linesep)
    pyperclip.copy(flow["user_code"])
    print(
        f"Access code '{flow['user_code']}' copied to clipboard. Pick a Microsoft account and sign in to OneDrive SyncEngine to begin download!",
        end=2 * os.linesep,
    )
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", "auth failed"))
    return app


def get_item(item_id, drive_id, token):
    # no $select - downloadUrl comes by default for file items
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, timeout=60)
    resp.raise_for_status()
    return resp.json()


def get_root_item(sharing_url, token):
    encoded = encode_sharing_url(sharing_url)
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{encoded}/driveItem",
        headers=headers,
        params={"$select": "id,name,parentReference"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def is_complete(dest_path, expected_size, expected_mtime):
    if not os.path.exists(dest_path):
        return False
    if os.path.getsize(dest_path) != expected_size:
        return False
    return abs(os.path.getmtime(dest_path) - expected_mtime) < 2  #  FAT32 limitation


def list_children(item_id, drive_id, token):
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"
    params = {"$select": "id,name,file,folder,size,lastModifiedDateTime"}
    headers = {"Authorization": f"Bearer {token}"}
    while url:
        resp = requests.get(url, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        yield from data.get("value", [])
        url = data.get("@odata.nextLink")
        params = None


def parse_mtime(iso_str):
    # OneDrive format: "2023-04-15T10:30:00Z"
    return datetime.datetime.fromisoformat(iso_str.replace("Z", "+00:00")).timestamp()


def refresh_token(app):
    # msal returns cached token if still valid, refreshes otherwise
    accounts = app.get_accounts()
    result = app.acquire_token_silent(scopes, account=accounts[0])
    if not result or "access_token" not in result:
        raise RuntimeError("token refresh failed - re-authenticate")
    return result["access_token"]


def stream_with_retry(url, dest_path, mtime, session, max_attempts=5):
    part_path = dest_path + ".part"
    for attempt in range(max_attempts):
        try:
            resume_from = os.path.getsize(part_path) if os.path.exists(part_path) else 0
            headers = {"Range": f"bytes={resume_from}-"} if resume_from else {}
            with session.get(url, headers=headers, stream=True, timeout=60) as r:
                r.raise_for_status()
                # server may ignore Range and return 200 with full content
                if resume_from and r.status_code != 206:
                    resume_from = 0
                mode = "ab" if resume_from else "wb"
                with open(part_path, mode) as f:
                    for chunk in r.iter_content(chunk_size=4_194_304):  # 4 MiB
                        f.write(chunk)
            os.replace(part_path, dest_path)
            os.utime(dest_path, (mtime, mtime))
            return
        except (
            requests.ConnectionError,
            requests.Timeout,
            requests.exceptions.ChunkedEncodingError,
        ) as e:
            err = e
        except requests.HTTPError as e:
            if e.response.status_code not in (429, 500, 502, 503, 504):
                raise
            err = e
        if attempt == max_attempts - 1:
            raise err
        wait = 2**attempt
        print(f"retry  {dest_path} in {wait}s ({err})")
        time.sleep(wait)


# access OneDrive data
with open("anyone.url", "r") as url_file:
    url_str = url_file.readline().strip()
app = get_access_token()

# download OneDrive data
print(f"FOLDER: {os.path.dirname(dest_root)}", end=2 * os.linesep, flush=True)
download_folder(url_str, app, dest_root)
print()
