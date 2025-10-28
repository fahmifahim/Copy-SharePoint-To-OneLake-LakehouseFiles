# SharePoint to OneLake (Lakehouse Files) — Recursive Folder Copy with MSAL & Graph
## Intro
This repository offers a Python-based solution to recursively copy the contents of a SharePoint folder—including all subfolders—into OneLake (Lakehouse Files). It uses MSAL for authentication and Microsoft Graph to access SharePoint data. While running, the notebook can also grant least-privilege per-site access (Sites.Selected) to another app, enabling future automations to safely access the same site.

You may go through see the code here:  
**sharepoint-to-onelake-recursive-copy.ipynb**

## Why copy files from SharePoint Online to Fabric OneLake (Lakehouse Files)?
Because analytics lives in OneLake. Teams drop PDFs, images, CSVs, and ad-hoc exports into SharePoint all the time—but your dashboards, notebooks, and pipelines run in Fabric. Pulling content from SharePoint → OneLake (/Files) lets you:

- Centralize data where Fabric engines can read it (Spark, SQL, Semantic models). OneLake is the “OneDrive for data,” automatically included with every Fabric tenant.  
- Automate ingestion on a schedule instead of manual uploads.  
- Keep least-privilege security using Sites.Selected so your app only sees the one site you grant.

## What you’ll build

```
App A (notebook) ── MSAL token ──> Microsoft Graph
                                     │
                                     ├─ Resolve siteId (/sites/{host}:/sites/{path})
                                     ├─ Ensure Sites.Selected (App B the least-privilege) grant on target site. 
                                     ├─ Locate 'Documents' drive
                                     ├─ Recursively list children of target folder and subfolder
                                     └─ Download each file (GET @microsoft.graph.downloadUrl)

                 ──> Write to Fabric Lakehouse /Files/<site>/<folder>/... (base64 via notebookutils/mssparkutils)
```

## Prerequisites
- A Fabric Lakehouse attached to the notebook (for /Files).
- Two Entra apps:
  - App A: the notebook’s caller (gets Graph token).
  - App B: the app you explicitly grant to the SharePoint site via Sites.Selected (least privilege).

## Quick preps for Microsoft Entra ID Apps registration

1. Register App A and App B  
- Entra admin center → App registrations → New registration. Note/copy the Application (client) ID and Directory (tenant) ID.  
[Microsoft Learn](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app)  
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/auth-register-app-v2)  

2. Create a client secret for App A (or use a certificate)  
  - App A → Certificates & secrets → New client secret. Note/copy the secret value (shown once).  
[Microsoft Learn](https://learn.microsoft.com/en-us/entra/identity-platform/howto-create-service-principal-portal)  

3. API permissions on App A and App B  
  - App A:  
    - Add Microsoft Graph → Application permission Site.Full Controll.All.  Click Grant admin consent.  
    - App A acts as the "admin account" who will authorize the next App Registration to access SharePoint.  

  - App B:  
    - Add SharePoint → Application permission Sites.Selected.  Click Grant admin consent.  

[Microsoft Learn](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/grant-admin-consent)


## Configure the notebook

In **Cell 1 — Parameters**, set:

```python
# ── App A: the app that this notebook uses to call Microsoft Graph (client-credential flow)
GRAPH_APP_CLIENT_ID       = "<App A - Client ID>"
GRAPH_APP_CLIENT_SECRET   = "<App A - Client Secret>"
GRAPH_TENANT_ID           = "<Tenant (Directory) ID>"

# ── SharePoint tenant & site (from your URL)
SP_TENANT_HOST            = "MySharePoint.sharepoint.com"
SP_SITE_PATH              = "MySites123"       # after /sites/
SP_TARGET_FOLDER_PATH     = "01_MyFolders"     # inside 'Documents' library; use "A/B/C" for deeper paths

# ── App B: the SECOND app you want to GRANT on this site (Sites.Selected target)
TARGET_APP_CLIENT_ID      = "<App B - Client ID>"
TARGET_APP_DISPLAY        = "MySite123-SharePoint-Access"

# ── If App A also needs to read/move files now, grant it too (helps avoid 403)
GRANT_CALLER_APP_TOO      = True

# ── Lakehouse Files destination
LAKEHOUSE_FILES_ROOT      = "Files"              # Fabric Lakehouse "Files" area (don’t change)
LAKEHOUSE_SITE_FOLDER     = SP_SITE_PATH         # keep per-site subfolder organization
```

```
※ Mapping example (target copy files from SharePoint):
https://MySharePoint.sharepoint.com/sites/MySites123/Shared Documents/01_MyFolders/image1.png
→
SP_TENANT_HOST="MySharePoint.sharepoint.com",
SP_SITE_PATH="MySites123",
SP_TARGET_FOLDER_PATH="01_MyFolders".
```

## How the code works (step-by-step)

1. Install MSAL  
`pip install msal` in the first cell.  

2. Get a Graph token (App A)  
`MSAL client credentials → acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])`  
[Microsoft Learn](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens)  

3. Resolve the site  
`GET /sites/{hostname}:/sites/{sitePath}` returns the siteId you’ll use for later calls.  
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/site-getbypath?view=graph-rest-1.0)  

4. Grant site access (Sites.Selected)  
- List existing site permissions: `GET /sites/{siteId}/permissions`.  
- If App B isn’t present, create a site permission: `POST /sites/{siteId}/permissions` with `"roles": ["write"]` and App B’s client id.  
- Optionally grant App A too so the notebook can copy files immediately.  
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/site-list-permissions?view=graph-rest-1.0)  

5. Locate the “Documents” drive  
`GET /sites/{siteId}/drives` → pick the documentLibrary named “Documents” (or the first library if localized).  
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/resources/sharepoint?view=graph-rest-1.0)  

6. Resolve the target folder  
`GET /drives/{driveId}/root:/{SP_TARGET_FOLDER_PATH}` to get the folder’s item id.  
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/driveitem-get?view=graph-rest-1.0)  

7. List files recursively  
Depth-first traversal using `GET /drives/{driveId}/items/{folderItemId}/children`, following `@odata.nextLink` for paging; when you hit a subfolder, recurse.  
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0)  

8. Download each file  
Use the file’s `@microsoft.graph.downloadUrl` to get bytes. (This is a short-lived pre-authenticated URL.)  
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/driveitem-get?view=graph-rest-1.0)  

9. Write to Lakehouse /Files  
- We mirror the SharePoint structure under:  
  `Files/<site>/<SP_TARGET_FOLDER_PATH>/<subfolders...>/<filename>`  
- `notebookutils.fs.mkdirs(...)` creates directories;  
- `mssparkutils.fs.put(...)` writes base64 text (because put is text-only).  
[Microsoft Learn](https://learn.microsoft.com/en-us/fabric/data-engineering/notebook-utilities)  
[Microsoft Learn](https://learn.microsoft.com/en-us/fabric/data-engineering/microsoft-spark-utilities)  

10. Logs  
For each file, the notebook prints a comment line showing the exact mapping:  
`# COPY: SP 'Documents/01.folder/.../file' -> Lakehouse 'Files/<site>/01.folder/.../file'`

## Run it

1. Open the notebook in Fabric and attach your Lakehouse.  
2. Edit Cell 1 — Parameters with your values.  
3. Run cells top to bottom.  
4. Validate in Lakehouse → Files that paths/files appear as expected.  

### Option to automate the run:  
**Operationalize:**  
- Add the notebook to a Fabric Data Pipeline with a schedule trigger, or  
- Use a Notebook schedule (depending on your workspace governance).  

Either way, you now have a repeatable pull from SharePoint to OneLake.

## Security notes
1. Prefer Sites.Selected over tenant-wide Sites.Read.All to minimize blast radius. It requires two steps: (a) add the permission to the app; (b) assign access to the specific site (this notebook does step b). 
2. Rotate client secrets regularly; consider certificates for stronger auth. 
3. Store secrets in Fabric workspace credentials / key vault when possible (avoid hard-coding in notebooks).

## References

- MSAL (Python) – acquire_token_for_client  
[Microsoft Learn](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens)

- Get SharePoint site by path (/sites/{hostname}:/{relative-path})  
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/site-getbypath?view=graph-rest-1.0)

- List folder children (/drives/{drive-id}/items/{item-id}/children)  
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0)

- DriveItem & download (/drives/{drive-id}/items/{item-id}, /content)  
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/driveitem-get?view=graph-rest-1.0)

- Create site permission (POST /sites/{site-id}/permissions)  
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/site-list-permissions?view=graph-rest-1.0)

- Selected permissions (Sites.Selected) overview  
[Microsoft Learn](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread)

- NotebookUtils / MSSparkUtils for Fabric (filesystem APIs)  
[Microsoft Learn](https://learn.microsoft.com/en-us/fabric/data-engineering/notebook-utilities)

