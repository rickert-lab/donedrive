# donedrive ☁️

_donedrive_ is a minimalistic tool to download OneDrive shares.

<img src="./assets/gui.png" alt="GUI with macOS">  

Microsoft imposes arbitrary [restrictions and limitations](https://support.microsoft.com/en-us/office/restrictions-and-limitations-in-onedrive-and-sharepoint-64883a5d-228e-48f5-b3d2-eb39e07630fa) in OneDrive and SharePoint. However, the installation of the [OneDrive client](https://www.microsoft.com/en-us/microsoft-365/onedrive/download) cannot be expected as a workaround for interrupted downloads or [missing files](https://learn.microsoft.com/en-us/answers/questions/1345845/how-to-download-a-zip-file-on-share-point-containi) in bulk downloads.

_donedrive_ offers a robust alternative to Microsoft's approach - effectively positioned between the OneDrive client and the SharePoint browser access. A simple OneDrive share link is sufficient for an authorized user to start the download of shared data: Client authentication is managed with [Microsoft's identity platform](https://learn.microsoft.com/en-us/graph/auth-v2-user?tabs=http) and data access is performed via [Microsoft's Graph API](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0&tabs=http#downloading-files-in-javascript-apps).

A couple of features that _donedrive_ offers:

* Concurrent data download to maximize transfer speed
* Automatic token refresh for downloads of very large files
* Resuming of interrupted file downloads during session
* Incremental downloads of missing or changed files

>[!CAUTION]
>Data integrity is not guaranteed.
