# donedrive ☁️

_donedrive_ is a minimalistic tool to download OneDrive shares.

<img src="./assets/gui.png" alt="GUI with macOS">

Microsoft imposes conservative [restrictions and limitations](https://support.microsoft.com/en-us/office/restrictions-and-limitations-in-onedrive-and-sharepoint-64883a5d-228e-48f5-b3d2-eb39e07630fa) in OneDrive and SharePoint. However, the installation of the [OneDrive client](https://www.microsoft.com/en-us/microsoft-365/onedrive/download) cannot be expected as the default workaround for [interrupted downloads](https://learn.microsoft.com/en-us/answers/questions/5204975/unable-to-download-large-files-14gb-from-sharepoin) of large files or for [missing files](https://learn.microsoft.com/en-us/answers/questions/1345845/how-to-download-a-zip-file-on-share-point-containi) in bulk downloads.

_donedrive_ offers a robust alternative to Microsoft's approach - effectively positioned between the powerful OneDrive client and the simple SharePoint browser access. A single OneDrive share link is sufficient to start the download of shared data: _donedrive_ follows the link to obtain [anonymous access](https://learn.microsoft.com/en-us/sharepoint/shareable-links-anyone-specific-people-organization#anyone-links) to the share and enumerates its contents through SharePoint's endpoints.

>[!IMPORTANT]
>User access could be delegated to applications like _donedrive_ with user authentication: However, application access must finally be authorized by every admin of every single SharePoint installation. Therefore, _donedrive_ only accepts share links without any user authentication and without any password:
>
><img src="./assets/share.png" alt="GUI with Anyone" width="50%">

A couple of features that _donedrive_ offers:

* Concurrent data download to maximize transfer speed
* Automatic download URL refresh for downloads of very large files
* Resuming of interrupted file downloads across sessions
* Incremental downloads of missing or changed files

>[!TIP]
>Incremental downloads are "baked" into _donedrive_. Simply download the same share to the same destination directory at a later point in time. The little icons following the file paths indicate the download status of individual files: Remote files that need to be downloaded in full indicated by a down arrow (↓), partial downloads by a double arrow (⇅). Downloaded files that have been validated from disk are indicated by a check mark (✓) - they don't need to be downloaded again.

Example output:
```terminal
==================================================================================
ROOT:  /Users/user/Desktop

Opened share link, starting up to 10 concurrent downloads:

DIR:   /test ✔
DIR:   /test/big ✔
FILE:  /test/big/bigger.bin [1.0 GiB] ⇅
DIR:   /test/mid ✔
FILE:  /test/big/big.bin [100.0 MiB] ✔
FILE:  /test/mid/mid.bin [4.0 MiB] ⭣
FILE:  /test/smol.bin [1.0 MiB] ✔

==================================================================================
DIRS:  3 ✔
FILES: 4 [1.1 GiB] ✔
==================================================================================

Download completed in 19.7 seconds with ~57.4 MiB/s.
```

>[!CAUTION]
>The integrity of the downloaded data is confirmed by matching the modification times, sizes, and hashes of the local copies with those of the remote copies. However, the calculation of the QuickXorHash values is computationally expensive and will take some time to complete - especially for large (> 1 GiB) files.
