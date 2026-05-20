#!/usr/bin/env python3

import base64
import datetime
import os
import threading
import time
import tkinter as tk
import webbrowser

from concurrent.futures import ThreadPoolExecutor
from tkinter import filedialog

import msal
import pyperclip
import requests


client_id = "ab9b8c07-8f02-4f72-87fa-80105867a763"  # OneDrive SyncEngine ID
scopes = ["https://graph.microsoft.com/Files.Read"]  # Microsoft Graph API
max_concurrent = 10  # max parallel downloads


def _download_objects(
    item_id, drive_id, app, session, dest_dir, dest_root, executor, futures
):
    os.makedirs(dest_dir, exist_ok=True)
    print(f"DIR:  {os.path.sep + os.path.relpath(dest_dir, dest_root)}")
    token = refresh_token(app)
    for item in list_objects(item_id, drive_id, token):
        dest_path = os.path.join(dest_dir, item["name"])
        if "folder" in item:
            _download_objects(
                item["id"],
                drive_id,
                app,
                session,
                dest_path,
                dest_root,
                executor,
                futures,
            )
        elif "file" in item:
            futures.append(
                executor.submit(
                    download_file, item, drive_id, dest_path, dest_root, app, session
                )
            )


def browse_folder(path_var, path_entry):
    folder_path = filedialog.askdirectory()
    if folder_path:
        path_var.set(folder_path)
        path_entry.xview_moveto(1.0)


def build_gui():
    root = tk.Tk()
    root.title("donedrive")

    url_var = tk.StringVar(
        value=18 * " " + "Copy & paste your OneDrive share link to download."
    )
    path_var = tk.StringVar(
        value="Browse to your download destination folder." + 26 * " "
    )

    url_frame = tk.Frame(root)
    url_frame.pack(fill="x", padx=10, pady=(10, 5))
    tk.Button(url_frame, text="Paste", command=lambda: paste_url(url_var)).pack(
        side="left"
    )
    tk.Entry(url_frame, textvariable=url_var, state="readonly", width=60).pack(
        side="left", fill="x", expand=True, padx=(5, 0)
    )

    path_frame = tk.Frame(root)
    path_frame.pack(fill="x", padx=10, pady=5)
    path_entry = tk.Entry(
        path_frame, textvariable=path_var, state="readonly", justify="right"
    )
    path_entry.pack(side="left", fill="x", expand=True)
    tk.Button(
        path_frame, text="Browse", command=lambda: browse_folder(path_var, path_entry)
    ).pack(side="left", padx=(5, 0))

    download_btn = tk.Button(root, text="Download", state="disabled")
    download_btn.config(command=lambda: start_download(url_var, path_var, download_btn))
    download_btn.pack(pady=10)

    footer = tk.Frame(root)
    footer.pack(fill="x", side="bottom", padx=10, pady=(0, 10))
    link = tk.Label(
        footer,
        text="Version 0.1 by Christian Rickert. ↗",
        fg="light blue",
        font=("TkDefaultFont", 9),
    )
    link.pack(side="right")
    link.bind(
        "<Button-1>",  # left-click
        lambda b: webbrowser.open("https://github.com/rickert-lab/donedrive"),
    )
    link.bind("<Enter>", lambda c: link.config(font=("TkDefaultFont", 9, "underline")))
    link.bind("<Leave>", lambda c: link.config(font=("TkDefaultFont", 9)))

    update = lambda *_: update_download_state(url_var, path_var, download_btn)
    url_var.trace_add("write", update)
    path_var.trace_add("write", update)

    root.update_idletasks()  # force geometry calculation before querying size
    root.minsize(root.winfo_width(), root.winfo_height())
    root.resizable(True, False)

    root.mainloop()


def byte_size(num_bytes):
    for unit in ("B", "KiB", "MiB", "GiB", "TiB"):
        if num_bytes < 1024:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.1f} PiB"


def download_file(item, drive_id, dest_path, dest_root, app, session):
    expected_size = item["size"]
    expected_mtime = parse_mtime(item["lastModifiedDateTime"])
    if is_complete(dest_path, expected_size, expected_mtime):
        print(
            f"FILE: {os.path.sep + os.path.relpath(dest_path, dest_root)} [skip]",
            flush=True,
        )
        return
    print(
        f"FILE: {os.path.sep + os.path.relpath(dest_path, dest_root)} [{byte_size(expected_size)}]",
        flush=True,
    )
    stream_with_retry(
        item, drive_id, dest_path, expected_mtime, expected_size, app, session
    )


def download_folder(sharing_url, app, dest_dir):
    token = refresh_token(app)
    root = get_root_item(sharing_url, token)
    drive_id = root["parentReference"]["driveId"]
    dest_dir = os.path.join(dest_dir, root["name"])
    futures = []
    with (
        requests.Session() as session,
        ThreadPoolExecutor(max_workers=max_concurrent) as executor,
    ):
        # match pool size to worker count so threads don't serialize on the pool
        adapter = requests.adapters.HTTPAdapter(
            pool_connections=max_concurrent, pool_maxsize=max_concurrent
        )
        session.mount("https://", adapter)
        _download_objects(
            root["id"],
            drive_id,
            app,
            session,
            dest_dir,
            os.path.dirname(dest_dir),  # dest_root
            executor,
            futures,
        )
    # surface exceptions from worker threads (with-block already waited for all)
    errors = []
    for f in futures:
        try:
            f.result()
        except Exception as e:
            errors.append(e)
    if errors:
        for e in errors:
            print(f"ERROR: {e}")
        raise errors[0]
    else:
        dirs, files, total = summarize_download(dest_dir)
        print(f"{os.linesep}{dirs} dirs, {files} files, {byte_size(total)}")
        print(f"Download complete.{os.linesep}")


def encode_sharing_url(url):
    # Microsoft Graph API requires base64url encoding with u! prefix
    encoded = base64.urlsafe_b64encode(url.encode()).rstrip(b"=").decode()
    return "u!" + encoded


def get_client_app():
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
        f"Access code '{flow['user_code']}' copied to the clipboard. Pick a Microsoft account and sign in to OneDrive SyncEngine to begin the download!",
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


def list_objects(item_id, drive_id, token):
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


def paste_url(url_var):
    url_var.set(pyperclip.paste())


def refresh_token(app):
    accounts = app.get_accounts()
    result = app.acquire_token_silent(scopes, account=accounts[0]) if accounts else None
    if not result or "access_token" not in result:
        raise RuntimeError("token refresh failed - re-authenticate")
    return result["access_token"]


def start_download(url_var, path_var, button):
    url = url_var.get().strip()
    dest = path_var.get().strip()
    if not url or not dest:
        return
    button.config(state="disabled")

    def worker():
        try:
            app = get_client_app()
            print(f"ROOT: {dest}", end=2 * os.linesep, flush=True)
            download_folder(url, app, dest)
            print()
        finally:
            # marshal Tk call back to the main thread
            button.after(0, lambda: button.config(state="normal"))

    threading.Thread(target=worker, daemon=True).start()


def stream_with_retry(
    item, drive_id, dest_path, mtime, expected_size, app, session, max_attempts=5
):
    part_path = dest_path + ".part"
    # discard any stale .part from a previous run - can't verify it matches this item
    if os.path.exists(part_path):
        try:
            os.remove(part_path)
        except OSError:
            pass
    url = None
    for attempt in range(max_attempts):
        try:
            if url is None:
                token = refresh_token(app)
                url = get_item(item["id"], drive_id, token)[
                    "@microsoft.graph.downloadUrl"
                ]
            resume_from = os.path.getsize(part_path) if os.path.exists(part_path) else 0
            headers = {"Range": f"bytes={resume_from}-"} if resume_from else {}
            with session.get(url, headers=headers, stream=True, timeout=60) as r:
                r.raise_for_status()
                if resume_from and r.status_code != 206:
                    resume_from = 0
                mode = "ab" if resume_from else "wb"
                with open(part_path, mode) as f:
                    for chunk in r.iter_content(chunk_size=4_194_304):
                        f.write(chunk)
            if os.path.getsize(part_path) != expected_size:
                actual_size = os.path.getsize(part_path)
                try:
                    os.remove(part_path)
                except OSError:
                    pass
                raise IOError(
                    f"ERROR: File size is {byte_size(actual_size)}, but expected {byte_size(expected_size)}."
                )
            os.replace(part_path, dest_path)
            os.utime(dest_path, (mtime, mtime))
            return
        except (
            requests.ConnectionError,
            requests.Timeout,
            requests.exceptions.ChunkedEncodingError,
            IOError,
        ) as e:
            err = e
            url = None  # force refresh next attempt
        except requests.HTTPError as e:
            if e.response.status_code not in (429, 500, 502, 503, 504):
                raise
            err = e
            url = None  # force refresh next attempt
        if attempt == max_attempts - 1:
            raise err
        wait = 2**attempt
        print(f"RETRY: {dest_path} in {wait}s ({err})")
        time.sleep(wait)


def summarize_download(dest_dir):
    dirs = files = total = 0
    for root, _dnames, fnames in os.walk(dest_dir):
        dirs += 1
        for fname in fnames:
            files += 1
            total += os.path.getsize(os.path.join(root, fname))
    return dirs, files, total


def update_download_state(url_var, path_var, button):
    url = url_var.get().strip().lower()
    dest = path_var.get().strip()
    valid_domains = ("1drv.ms", "onedrive.live.com", "sharepoint.com", "onedrive.com")
    url_ok = url.startswith("https://") and any(d in url for d in valid_domains)
    valid = url_ok and os.path.isdir(dest)
    button.config(state="normal" if valid else "disabled")


# start GUI
build_gui()
