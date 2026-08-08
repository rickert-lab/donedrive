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
Date:       2026-08-07
Version:    0.2
"""

import base64
import datetime
import os
import stat
import threading
import time
import tkinter as tk
import webbrowser

from concurrent.futures import CancelledError, ThreadPoolExecutor
from tkinter import filedialog
from urllib.parse import urlparse

import numpy as np

# os.environ["GTK_MODULES"] = ""  # do not autload 'atk-bridge' (LINUX)
import pyperclip
import requests


max_concurrent = 10  # max concurrent downloads
cancel_event = threading.Event()  # set on window close to abort in-flight downloads


class DownloadCancelled(Exception):
    """Raised in worker threads when the user closes the window."""


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

        # vectorized strided XOR-reduction: reduced[i] = XOR of bytes at i, i+160, i+320, ...
        # lay the chunk out as a 2D grid 160 columns wide, then XOR-reduce each column.
        buf = np.frombuffer(array, dtype=np.uint8, count=cbSize, offset=ibStart)
        n_full = cbSize // self.WidthInBits  # number of complete 160-byte rows
        remainder = (
            cbSize % self.WidthInBits
        )  # trailing bytes that don't form a full row
        reduced = np.zeros(self.WidthInBits, dtype=np.uint64)
        if n_full:
            # reshape into (n_full, 160) and collapse rows -> one byte per column
            reduced[:] = np.bitwise_xor.reduce(
                buf[: n_full * self.WidthInBits].reshape(n_full, self.WidthInBits),
                axis=0,
            ).astype(np.uint64)
        if remainder:
            # fold the leftover bytes into the matching columns of the reduction
            reduced[:remainder] ^= buf[n_full * self.WidthInBits :].astype(np.uint64)

        # from here on it's the original per-column placement into the 160-bit state,
        # but driven by the precomputed `reduced` array instead of an inner byte loop.
        currentShift = self._shiftSoFar

        # the bitvector where we'll start xoring
        vectorArrayIndex = currentShift // 64

        # the position within the bit vector at which we begin xoring
        vectorOffset = currentShift % 64
        iterations = min(cbSize, self.WidthInBits)

        for i in range(iterations):
            isLastCell = vectorArrayIndex == len(self._data) - 1
            bitsInVectorCell = self.BitsInLastCell if isLastCell else 64
            xoredByte = int(reduced[i])  # precomputed column XOR

            # there's at least 2 bitvectors before we reach the end of the array
            if vectorOffset <= bitsInVectorCell - 8:
                # whole byte fits inside the current 64-bit cell - no mask needed
                self._data[vectorArrayIndex] ^= xoredByte << vectorOffset
            else:
                # byte straddles the cell boundary - split across two cells, wrapping at the end
                index1 = vectorArrayIndex
                index2 = 0 if isLastCell else (vectorArrayIndex + 1)
                low = bitsInVectorCell - vectorOffset
                self._data[index1] = (
                    self._data[index1] ^ (xoredByte << vectorOffset)
                ) & self._MASK64
                self._data[index2] ^= xoredByte >> low

            vectorOffset += self.Shift
            while vectorOffset >= bitsInVectorCell:
                vectorArrayIndex = 0 if isLastCell else vectorArrayIndex + 1
                vectorOffset -= bitsInVectorCell

        # update the starting position in a circular shift pattern
        self._shiftSoFar = (
            self._shiftSoFar + self.Shift * (cbSize % self.WidthInBits)
        ) % self.WidthInBits

        self._lengthSoFar += cbSize

    def digest(self):
        # create a byte array big enough to hold all our data
        rgb = bytearray((self.WidthInBits - 1) // 8 + 1)

        # block copy all our bitvectors to this byte array
        for i in range(len(self._data) - 1):
            rgb[i * 8 : i * 8 + 8] = self._data[i].to_bytes(8, "little")

        last = len(self._data) - 1
        tail_len = len(rgb) - last * 8
        rgb[last * 8 :] = (self._data[last] & self._MASK64).to_bytes(8, "little")[
            :tail_len
        ]

        # XOR the file length with the least significant bits
        # expected value is 8-bytes in length in little-endian format
        lengthBytes = (self._lengthSoFar & self._MASK64).to_bytes(8, "little")
        for i in range(len(lengthBytes)):
            rgb[(self.WidthInBits // 8) - len(lengthBytes) + i] ^= lengthBytes[i]

        return bytes(rgb)


def _download_objects(
    item_id, drive_id, api, session, dest_dir, dest_root, executor, futures
):
    if os.path.isdir(dest_dir):
        print(
            f"DIR:   {os.path.sep + os.path.relpath(dest_dir, dest_root)} ✓", flush=True
        )
    else:
        os.makedirs(dest_dir, exist_ok=True)
        print(
            f"DIR:   {os.path.sep + os.path.relpath(dest_dir, dest_root)} ⭣", flush=True
        )
    for item in list_objects(item_id, drive_id, api, session):
        if cancel_event.is_set():
            break
        dest_path = safe_join(dest_dir, item["name"])
        if "folder" in item:  # recurse
            _download_objects(
                item["id"],
                drive_id,
                api,
                session,
                dest_path,
                dest_root,
                executor,
                futures,
            )
        elif "file" in item:  # download
            futures.append(
                executor.submit(
                    download_file, item, drive_id, dest_path, dest_root, api, session
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
        side="left", padx=(0, 5), pady=0
    )
    url_entry = tk.Entry(
        url_frame,
        textvariable=url_var,
        state="readonly",
        justify="left",
        width=60,
        foreground="white",  # fix font color for dark/ligth mode switch (macos)
        readonlybackground="gray30",  # color changes dynamically (macos)
    )
    url_entry.pack(side="left", fill="x", expand=True, padx=0, pady=(4, 0))

    # download destination
    dir_frame = tk.Frame(root)
    dir_frame.pack(fill="x", padx=10, pady=(5, 10))
    dir_entry = tk.Entry(
        dir_frame,
        textvariable=dir_var,
        state="readonly",
        justify="right",
        width=60,
        foreground="white",
        readonlybackground="gray30",
    )
    dir_entry.pack(side="left", fill="x", expand=True, padx=0, pady=(4, 0))
    tk.Button(
        dir_frame, text="Browse", command=lambda: browse_folder(dir_var, dir_entry)
    ).pack(side="left", padx=(5, 0), pady=0)

    download_btn = tk.Button(root, text="Download", state="disabled")
    download_btn.config(command=lambda: start_download(url_var, dir_var, download_btn))
    download_btn.pack(padx=0, pady=(5, 0))

    footer = tk.Frame(root)
    footer.pack(fill="x", side="bottom", padx=5, pady=5)
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
        text="Version 0.2 by Christian Rickert. ↗",
        fg="dodger blue",
        font=("TkDefaultFont", 10),
    )
    author.pack(side="right")
    author.bind(
        "<Button-1>",  # left-click
        lambda _e: webbrowser.open("https://github.com/rickert-lab/donedrive"),
    )
    author.bind(
        "<Enter>", lambda _e: author.config(font=("TkDefaultFont", 10, "underline"))
    )
    author.bind("<Leave>", lambda _e: author.config(font=("TkDefaultFont", 10)))

    # keep track of download status
    def update(*_):
        update_download_state(url_var, dir_var, download_btn)

    url_var.trace_add("write", update)
    dir_var.trace_add("write", update)

    # abort in-flight downloads instead of blocking on them at exit
    def on_close():
        cancel_event.set()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    # limit resizing to width
    root.update_idletasks()  # force geometry calculation before querying size
    root.minsize(root.winfo_width(), root.winfo_height())
    root.resizable(True, False)

    # run GUI
    root.mainloop()


def download_file(item, drive_id, dest_path, dest_root, api, session):
    if cancel_event.is_set():  # future started before shutdown reached it
        return
    expected_size = item["size"]
    expected_mtime = parse_mtime(item["lastModifiedDateTime"])
    expected_hash = item.get("file", {}).get("hashes", {}).get("quickXorHash")
    if is_complete(dest_path, expected_size, expected_mtime, expected_hash):
        print(
            f"FILE:  {os.path.sep + os.path.relpath(dest_path, dest_root)} [{format_size(expected_size)}] ✓",
            flush=True,
        )
        return
    print(
        f"FILE:  {os.path.sep + os.path.relpath(dest_path, dest_root)} [{format_size(expected_size)}] ⭣",
        flush=True,
    )
    stream_with_retry(
        item,
        drive_id,
        dest_path,
        dest_root,
        expected_mtime,
        expected_size,
        expected_hash,
        api,
        session,
    )


def download_folder(sharing_url, dest_dir):
    session, api = get_anonymous_session(sharing_url)
    root = get_root_item(sharing_url, api, session)
    drive_id = root["parentReference"]["driveId"]
    dest_dir = safe_join(dest_dir, root["name"])
    futures = []
    with session, ThreadPoolExecutor(max_workers=max_concurrent) as executor:
        try:
            _download_objects(
                root["id"],
                drive_id,
                api,
                session,
                dest_dir,
                os.path.dirname(dest_dir),  # dest_root
                executor,
                futures,
            )
        finally:
            if cancel_event.is_set():
                executor.shutdown(cancel_futures=True)  # drop queued downloads
    # surface exceptions from worker threads (with-block already waited for all)
    errors = []
    for f in futures:
        try:
            f.result()
        except (CancelledError, DownloadCancelled):  # window closed
            pass
        except Exception as e:
            errors.append(e)
    if errors:
        for e in errors:
            print(e, flush=True)
        raise errors[0]
    if cancel_event.is_set():
        raise DownloadCancelled
    dirs, files, total = summarize_download(dest_dir)
    return (dirs, files, total)


def encode_sharing_url(url):
    # Microsoft Graph API requires base64url encoding with u! prefix
    encoded = base64.urlsafe_b64encode(url.encode()).rstrip(b"=").decode()
    return "u!" + encoded


def format_duration(seconds):
    if seconds < 60:
        return f"{seconds:.1f} seconds"
    h, r = divmod(round(seconds), 3_600)
    m, s = divmod(r, 60)
    if h:
        return f"{h} hours, {m} minutes, and {s} seconds"
    if m:
        return f"{m} minutes and {s} seconds"
    return f"{s} seconds"


def format_size(num_bytes):
    for unit in ("B", "KiB", "MiB", "GiB", "TiB"):
        if round(num_bytes, 1) < 1024:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.1f} PiB"


def get_anonymous_session(sharing_url):
    # follow the share link so SharePoint sets its anonymous access cookie
    session = requests.Session()
    session.headers["User-Agent"] = "Mozilla/5.0"  # SPO rejects the default UA
    # match pool size to worker count so threads don't serialize on the pool
    adapter = requests.adapters.HTTPAdapter(
        pool_connections=max_concurrent, pool_maxsize=max_concurrent
    )
    session.mount("https://", adapter)
    with session.get(sharing_url, allow_redirects=True, stream=True, timeout=30) as r:
        r.raise_for_status()
        host = urlparse(r.url).netloc  # resolves short links (1drv.ms) as well
    print(
        f"Opened share link, starting up to {max_concurrent} concurrent downloads:{os.linesep}",
        flush=True,
    )
    return session, f"https://{host}/_api/v2.0"


def get_download_url(item_id, drive_id, api, session):
    # SPO v2.0 names it "@content.downloadUrl"; Graph uses "@microsoft.graph.downloadUrl"
    item = spo_get(session, f"{api}/drives/{drive_id}/items/{item_id}", timeout=60)
    return item.get("@content.downloadUrl") or item["@microsoft.graph.downloadUrl"]


def get_root_item(sharing_url, api, session):
    encoded = encode_sharing_url(sharing_url)
    try:
        return spo_get(
            session,
            f"{api}/shares/{encoded}/driveItem",
            params={"$select": "id,name,parentReference"},
        )
    except requests.HTTPError as e:
        if e.response.status_code not in (401, 403):
            raise
        raise RuntimeError(
            os.linesep
            + "NOTE: The share link requires a sign-in. Ask the sender to share "
            'the folder with the "Anyone with the link" option instead: Shares '
            "addressed to a specific person cannot be accessed by this program."
        ) from e


def is_complete(dest_path, expected_size, expected_mtime, expected_hash):
    if not os.path.exists(dest_path):  # fast
        return False
    if abs(os.path.getmtime(dest_path) - expected_mtime) > 2:  # 2 s FAT32 resolution
        return False
    if os.path.getsize(dest_path) != expected_size:  # slower
        return False
    if expected_hash and quickxorhash_file(dest_path) != expected_hash:  # slowest
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


def list_objects(item_id, drive_id, api, session):
    url = f"{api}/drives/{drive_id}/items/{item_id}/children"
    params = {"$select": "id,name,file,folder,size,lastModifiedDateTime"}
    while url:
        data = spo_get(session, url, params=params)
        yield from data.get("value", [])
        url = data.get("@odata.nextLink")
        params = None


def parse_mtime(iso_str):
    # OneDrive format: "2023-04-15T10:30:00Z"
    return datetime.datetime.fromisoformat(iso_str.replace("Z", "+00:00")).timestamp()


def paste_url(url_var):
    url_var.set(pyperclip.paste())


def quickxorhash_file(path, chunk_size=16_777_216):  # 16 MiB
    h = QuickXorHash()
    with open(path, "rb") as f:
        while chunk := f.read(chunk_size):
            h.update(chunk)
    return base64.b64encode(h.digest()).decode()


def safe_join(dest_dir, name):
    dest_path = os.path.join(dest_dir, sanitize_name(name))
    # belt and braces: the joined path must be a direct child of dest_dir
    if os.path.dirname(os.path.abspath(dest_path)) != os.path.abspath(dest_dir):
        raise ValueError(f"ERROR: item escapes the destination directory: {name!r}")
    return dest_path


def sanitize_name(name):
    # reject traversal: a name must be one ordinary component, not a path
    # (os.path.basename ignores backslashes outside Windows, so test it here)
    if name in ("", ".", "..") or os.path.basename(name) != name or "\\" in name:
        raise ValueError(f"ERROR: unsafe item name: {name!r}")
    clean = "".join("_" if c in '<>:"/\\|?*' or ord(c) < 32 else c for c in name)
    clean = clean.rstrip(
        ". "
    )  # Windows drops these silently, breaking size/hash checks
    if not clean:
        raise ValueError(f"ERROR: item name is empty after sanitizing: {name!r}")
    return clean


def spo_get(session, url, params=None, timeout=30):
    resp = session.get(url, params=params, timeout=timeout)
    resp.raise_for_status()
    return resp.json()


def start_download(url_var, dir_var, button):
    url = url_var.get().strip()
    dest = dir_var.get().strip()
    if not url or not dest:
        return
    print(f"{os.linesep}{80 * '='}")
    print(f"ROOT:  {dest}{os.linesep}", flush=True)
    button.config(state="disabled")

    def worker():
        try:
            start = time.time()
            (dirs, files, total) = download_folder(url, dest)
            stop = time.time()
            duration = stop - start
            with_rate = (
                f" with ~{format_size(total / duration)}/s" if duration > 0 else ""
            )
            print(f"{os.linesep}{80 * '='}")
            print(f"DIRS:  {dirs} ✓")
            print(f"FILES: {files} [{format_size(total)}] ✓")
            print(f"{80 * '='}{os.linesep}")
            print(
                f"Download completed in {format_duration(duration)}{with_rate}.",
                flush=True,
            )
        except DownloadCancelled:
            print(f"NOTE: Download cancelled.{os.linesep}", flush=True)
        finally:
            if not cancel_event.is_set():  # widget is gone after root.destroy()
                # marshal Tk call back to the main thread
                button.after(0, lambda: button.config(state="normal"))

    threading.Thread(target=worker, daemon=True).start()


def stream_with_retry(
    item,
    drive_id,
    dest_path,
    dest_root,
    mtime,
    expected_size,
    expected_hash,
    api,
    session,
    max_attempts=5,
):
    part_path = dest_path + ".part"
    # discard partial files from previous run
    if os.path.exists(part_path):
        os.remove(part_path)
    # start download
    url = None
    for attempt in range(max_attempts):
        try:
            if url is None:
                url = get_download_url(item["id"], drive_id, api, session)
            resume_from = os.path.getsize(part_path) if os.path.exists(part_path) else 0
            headers = {"Range": f"bytes={resume_from}-"} if resume_from else {}
            with session.get(url, headers=headers, stream=True, timeout=60) as r:
                r.raise_for_status()
                if resume_from and r.status_code != 206:
                    resume_from = 0
                mode = "ab" if resume_from else "wb"
                with open(part_path, mode) as f:
                    for chunk in r.iter_content(chunk_size=4_194_304):
                        if cancel_event.is_set():
                            raise DownloadCancelled
                        f.write(chunk)
            # compare file size
            if os.path.getsize(part_path) != expected_size:
                actual_size = os.path.getsize(part_path)
                try:
                    os.remove(part_path)
                except OSError:
                    pass
                raise IOError(
                    f"ERROR: {os.path.sep + os.path.relpath(part_path, dest_root)} [{format_size(actual_size)}?] ↑"
                )
            # compare file hash
            if expected_hash:  # recent uploads might fail
                actual_hash = quickxorhash_file(part_path)
                if actual_hash != expected_hash:
                    actual_size = os.path.getsize(part_path)
                    try:
                        os.remove(part_path)
                    except OSError:
                        pass
                    raise IOError(
                        f"ERROR: {os.path.sep + os.path.relpath(part_path, dest_root)} [{format_size(actual_size)}!] ↑"
                    )
            # finalize download
            os.utime(part_path, (mtime, mtime))  # set access time, modification time
            os.replace(part_path, dest_path)  # move file into final path
            return
        except requests.HTTPError as e:
            if e.response.status_code not in (429, 500, 502, 503, 504):
                raise
            err = e
            url = None
        except (
            requests.ConnectionError,
            requests.Timeout,
            requests.exceptions.ChunkedEncodingError,
            IOError,
        ) as e:
            err = e
            url = None  # force refresh next attempt
        if attempt == max_attempts - 1:
            raise err
        wait = 2**attempt
        if cancel_event.wait(wait):  # interruptible backoff
            raise DownloadCancelled


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
    url = url_var.get().strip()
    dest = dir_var.get().strip()
    valid_domains = ("1drv.ms", "onedrive.live.com", "sharepoint.com", "onedrive.com")
    try:
        netloc = urlparse(url).netloc.lower()
    except ValueError:
        netloc = ""
    url_ok = url.lower().startswith("https://") and any(
        netloc == d or netloc.endswith("." + d) for d in valid_domains
    )
    dir_ok = os.path.isdir(dest) and os.access(dest, os.W_OK)
    button.config(state="normal" if (url_ok and dir_ok) else "disabled")


# start GUI
build_gui()
