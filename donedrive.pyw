#!/usr/bin/env python3

"""
donedrive - The download assistant for Microsoft OneDrive.
Copyright (C) 2026 The Regents of the University of Colorado

This program is free software: you can redistribute it and/or modify it under
the terms of the GNU General Public License as published by the Free Software
Foundation, version 3.

This program is distributed in the hope that it will be useful, but WITHOUT ANY
WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A
PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with
this program. If not, see <https://www.gnu.org/licenses/>.

Author:     Christian Rickert <christian.rickert@cuanschutz.edu>
Date:       2026-05-19
Version:    0.1
"""

import base64
import datetime
import os
import stat
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


class QuickXorHash:
    """
    See: https://learn.microsoft.com/en-us/onedrive/developer/code-snippets/quickxorhash?view=odsp-graph-online
    """

    BitsInLastCell = 32
    Shift = 11
    Threshold = 600
    WidthInBits = 160

    _MASK64 = 0xFFFFFFFFFFFFFFFF

    def __init__(self):
        self.Initialize()

    def Initialize(self):
        self._data = [0] * ((self.WidthInBits - 1) // 64 + 1)
        self._shiftSoFar = 0
        self._lengthSoFar = 0

    def update(self, array, ibStart=0, cbSize=None):
        if cbSize is None:
            cbSize = len(array) - ibStart

        currentShift = self._shiftSoFar

        # The bitvector where we'll start xoring
        vectorArrayIndex = currentShift // 64

        # The position within the bit vector at which we begin xoring
        vectorOffset = currentShift % 64
        iterations = min(cbSize, self.WidthInBits)

        for i in range(iterations):
            isLastCell = vectorArrayIndex == len(self._data) - 1
            bitsInVectorCell = self.BitsInLastCell if isLastCell else 64

            # There's at least 2 bitvectors before we reach the end of the array
            if vectorOffset <= bitsInVectorCell - 8:
                for j in range(ibStart + i, cbSize + ibStart, self.WidthInBits):
                    self._data[vectorArrayIndex] ^= array[j] << vectorOffset
                self._data[vectorArrayIndex] &= self._MASK64
            else:
                index1 = vectorArrayIndex
                index2 = 0 if isLastCell else (vectorArrayIndex + 1)
                low = bitsInVectorCell - vectorOffset

                xoredByte = 0
                for j in range(ibStart + i, cbSize + ibStart, self.WidthInBits):
                    xoredByte ^= array[j]
                self._data[index1] = (
                    self._data[index1] ^ (xoredByte << vectorOffset)
                ) & self._MASK64
                self._data[index2] ^= xoredByte >> low

            vectorOffset += self.Shift
            while vectorOffset >= bitsInVectorCell:
                vectorArrayIndex = 0 if isLastCell else vectorArrayIndex + 1
                vectorOffset -= bitsInVectorCell

        # Update the starting position in a circular shift pattern
        self._shiftSoFar = (
            self._shiftSoFar + self.Shift * (cbSize % self.WidthInBits)
        ) % self.WidthInBits

        self._lengthSoFar += cbSize

    def digest(self):
        # Create a byte array big enough to hold all our data
        rgb = bytearray((self.WidthInBits - 1) // 8 + 1)

        # Block copy all our bitvectors to this byte array
        for i in range(len(self._data) - 1):
            rgb[i * 8 : i * 8 + 8] = self._data[i].to_bytes(8, "little")

        last = len(self._data) - 1
        tail_len = len(rgb) - last * 8
        rgb[last * 8 :] = (self._data[last] & self._MASK64).to_bytes(8, "little")[
            :tail_len
        ]

        # XOR the file length with the least significant bits
        # Expected value is 8-bytes in length in little-endian format
        lengthBytes = (self._lengthSoFar & self._MASK64).to_bytes(8, "little")
        for i in range(len(lengthBytes)):
            rgb[(self.WidthInBits // 8) - len(lengthBytes) + i] ^= lengthBytes[i]

        return bytes(rgb)


def _download_objects(
    item_id, drive_id, app, session, dest_dir, dest_root, executor, futures
):
    if os.path.isdir(dest_dir):
        print(
            f"DIR:   {os.path.sep + os.path.relpath(dest_dir, dest_root)} ⭥", flush=True
        )
    else:
        os.makedirs(dest_dir, exist_ok=True)
        print(
            f"DIR:   {os.path.sep + os.path.relpath(dest_dir, dest_root)} ⭣", flush=True
        )
    token = refresh_token(app)
    for item in list_objects(item_id, drive_id, token):
        dest_path = os.path.join(dest_dir, item["name"])
        if "folder" in item:  # recurse
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
        elif "file" in item:  # download
            futures.append(
                executor.submit(
                    download_file, item, drive_id, dest_path, dest_root, app, session
                )
            )


def browse_folder(dir_var, dir_entry):
    folder_path = filedialog.askdirectory()
    if folder_path:
        dir_var.set(folder_path)
        dir_entry.xview_moveto(1.0)


def build_gui():
    root = tk.Tk()
    root.title("donedrive")

    # OneDrive share link
    url_var = tk.StringVar(value="[Paste the OneDrive share link for your download]")
    dir_var = tk.StringVar(value="[Select the destination directory for your download]")

    url_frame = tk.Frame(root)
    url_frame.pack(fill="x", padx=10, pady=(10, 5))
    tk.Button(url_frame, text="Paste", command=lambda: paste_url(url_var)).pack(
        side="left", padx=(5, 0), pady=(0, 5)
    )
    url_entry = tk.Entry(
        url_frame,
        textvariable=url_var,
        state="readonly",
        justify="left",
        width=60,
        readonlybackground="gray30",
    )
    url_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))

    # download destination
    dir_frame = tk.Frame(root)
    dir_frame.pack(fill="x", padx=10, pady=5)
    dir_entry = tk.Entry(
        dir_frame,
        textvariable=dir_var,
        state="readonly",
        justify="right",
        width=60,
        readonlybackground="gray30",
    )
    dir_entry.pack(side="left", fill="x", expand=True)
    tk.Button(
        dir_frame, text="Browse", command=lambda: browse_folder(dir_var, dir_entry)
    ).pack(side="left", padx=(5, 0), pady=(0, 5))

    download_btn = tk.Button(root, text="Download", state="disabled")
    download_btn.config(command=lambda: start_download(url_var, dir_var, download_btn))
    download_btn.pack(pady=10)

    footer = tk.Frame(root)
    footer.pack(fill="x", side="bottom", padx=10, pady=(0, 10))
    # OneDrive status
    status = tk.Label(
        footer,
        text="Microsoft OneDrive Status. ↗",
        fg="dark gray",
        font=("TkDefaultFont", 10),
    )
    status.pack(side="left")
    status.bind(
        "<Button-1>",  # left-click
        lambda _b: webbrowser.open(
            "https://www.aguidetocloud.com/service-health/?service=Microsoft+OneDrive&status=active"
        ),
    )
    status.bind(
        "<Enter>", lambda _c: status.config(font=("TkDefaultFont", 10, "underline"))
    )
    status.bind("<Leave>", lambda _c: status.config(font=("TkDefaultFont", 10)))
    # author
    author = tk.Label(
        footer,
        text="Version 0.1 by Christian Rickert. ↗",
        fg="dodger blue",
        font=("TkDefaultFont", 10),
    )
    author.pack(side="right")
    author.bind(
        "<Button-1>",  # left-click
        lambda _b: webbrowser.open("https://github.com/rickert-lab/donedrive"),
    )
    author.bind(
        "<Enter>", lambda _c: author.config(font=("TkDefaultFont", 10, "underline"))
    )
    author.bind("<Leave>", lambda _c: author.config(font=("TkDefaultFont", 10)))

    # keep track of download status
    update = lambda *_: update_download_state(url_var, dir_var, download_btn)
    url_var.trace_add("write", update)
    dir_var.trace_add("write", update)

    # limit resizing to width
    root.update_idletasks()  # force geometry calculation before querying size
    root.minsize(root.winfo_width(), root.winfo_height())
    root.resizable(True, False)

    # run GUI
    root.mainloop()


def byte_size(num_bytes):
    for unit in ("B", "KiB", "MiB", "GiB", "TiB"):
        if round(num_bytes, 1) < 1024:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.1f} PiB"


def download_file(item, drive_id, dest_path, dest_root, app, session):
    expected_size = item["size"]
    expected_mtime = parse_mtime(item["lastModifiedDateTime"])
    expected_hash = item.get("file", {}).get("hashes", {}).get("quickXorHash")
    if is_complete(dest_path, expected_size, expected_mtime, expected_hash):
        print(
            f"FILE:  {os.path.sep + os.path.relpath(dest_path, dest_root)} [{byte_size(expected_size)}] ⭥",
            flush=True,
        )
        return
    print(
        f"FILE:  {os.path.sep + os.path.relpath(dest_path, dest_root)} [{byte_size(expected_size)}] ⭣",
        flush=True,
    )
    stream_with_retry(
        item,
        drive_id,
        dest_path,
        expected_mtime,
        expected_size,
        expected_hash,
        app,
        session,
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
        print(f"{os.linesep}{80 * '='}")
        print(f"DIRS:  {dirs}")
        print(f"FILES: {files} [{byte_size(total)}]")
        print(f"{80 * '='}{os.linesep}")
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
    print(f'{os.linesep}"{flow["message"]}" - Microsoft', end=2 * os.linesep)
    pyperclip.copy(flow["user_code"])
    print(
        f"Access code '{flow['user_code']}' copied to the clipboard. Paste the "
        "code into the device authenticator. Then pick your Microsoft account "
        "and sign in to the OneDrive SyncEngine to begin your download!",
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


def is_complete(dest_path, expected_size, expected_mtime, expected_hash):
    if not os.path.exists(dest_path):
        return False
    if os.path.getsize(dest_path) != expected_size:
        return False
    if abs(os.path.getmtime(dest_path) - expected_mtime) >= 2:  # FAT32 limitation
        return False
    if expected_hash and quickxorhash_file(dest_path) != expected_hash:
        return False
    return True


def is_hidden(path):
    if os.path.basename(path).startswith("."):  # LINUX, macOS
        return True
    if os.name == "nt":  # Windows
        try:
            return bool(os.stat(path).st_file_attributes & stat.FILE_ATTRIBUTE_HIDDEN)
        except OSError:
            return False
    return False


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


def quickxorhash_file(path, chunk_size=4_194_304):
    h = QuickXorHash()
    with open(path, "rb") as f:
        while chunk := f.read(chunk_size):
            h.update(chunk)
    return base64.b64encode(h.digest()).decode()


def refresh_token(app):
    accounts = app.get_accounts()
    result = app.acquire_token_silent(scopes, account=accounts[0]) if accounts else None
    if not result or "access_token" not in result:
        raise RuntimeError("token refresh failed - re-authenticate")
    return result["access_token"]


def start_download(url_var, dir_var, button):
    url = url_var.get().strip()
    dest = dir_var.get().strip()
    if not url or not dest:
        return
    button.config(state="disabled")

    def worker():
        try:
            app = get_client_app()
            print(f"ROOT:  {dest}", end=2 * os.linesep, flush=True)
            download_folder(url, app, dest)
            print()
        finally:
            # marshal Tk call back to the main thread
            button.after(0, lambda: button.config(state="normal"))

    threading.Thread(target=worker, daemon=True).start()


def stream_with_retry(
    item,
    drive_id,
    dest_path,
    mtime,
    expected_size,
    expected_hash,
    app,
    session,
    max_attempts=5,
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
            if expected_hash:
                actual_hash = quickxorhash_file(part_path)
                if actual_hash != expected_hash:
                    try:
                        os.remove(part_path)
                    except OSError:
                        pass
                    raise IOError(
                        f"ERROR: File hash is {actual_hash}, but expected {expected_hash}."
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
    for root, dnames, fnames in os.walk(dest_dir):
        # prune hidden dirs in place so os.walk doesn't descend into them
        dnames[:] = [d for d in dnames if not is_hidden(os.path.join(root, d))]
        dirs += 1
        for fname in fnames:
            if fname.endswith(".part"):  # temp file
                continue
            fpath = os.path.join(root, fname)
            if is_hidden(fpath):
                continue
            files += 1
            total += os.path.getsize(fpath)
    return dirs, files, total


def update_download_state(url_var, dir_var, button):
    url = url_var.get().strip().lower()
    dest = dir_var.get().strip()
    valid_domains = ("1drv.ms", "onedrive.live.com", "sharepoint.com", "onedrive.com")
    url_ok = url.startswith("https://") and any(d in url for d in valid_domains)
    dir_ok = os.path.isdir(dest) and os.access(dest, os.W_OK)
    button.config(state="normal" if (url_ok and dir_ok) else "disabled")


# start GUI
build_gui()
