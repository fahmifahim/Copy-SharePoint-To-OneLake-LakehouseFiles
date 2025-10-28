## 0) Install dependency (once per session)

%pip install msal --quiet



## 1) Parameters (edit here)
# ── App A: the app that this notebook uses to call Microsoft Graph (client-credential flow)
GRAPH_APP_CLIENT_ID   = "<App A - Client ID>"
GRAPH_APP_CLIENT_SECRET = "<App A - Client Secret>"
GRAPH_TENANT_ID       = "<Tenant (Directory) ID>"

# ── SharePoint tenant & site (from your URL)
SP_TENANT_HOST        = "mngenvmcap535692.sharepoint.com"  # e.g., contoso.sharepoint.com
SP_SITE_PATH          = "sharepoint-dzlab2"                # after /sites/
SP_TARGET_FOLDER_PATH = "01.folder"                        # inside the default library (Documents). Use "A/B/C" for deeper paths.

# ── App B: the SECOND app you want to GRANT on this site (Sites.Selected target)
TARGET_APP_CLIENT_ID  = "<App B - Client ID>"
TARGET_APP_DISPLAY    = "dzlab2-SharePoint-Access"           # label only (for readability in audit)

# ── If App A also needs to read/move files now, grant it too (helps avoid 403)
GRANT_CALLER_APP_TOO  = True

# ── Lakehouse Files destination
LAKEHOUSE_FILES_ROOT  = "Files"              # Fabric Lakehouse "Files" area (don’t change)
LAKEHOUSE_SITE_FOLDER = SP_SITE_PATH         # keep per-site subfolder organization



## 2) Get a Graph token (client credentials with MSAL)

from msal import ConfidentialClientApplication

AUTHORITY = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}"
SCOPES    = ["https://graph.microsoft.com/.default"]

print("🔐 Getting Microsoft Graph token (client credentials)…")
msal_app = ConfidentialClientApplication(
    client_id=GRAPH_APP_CLIENT_ID,
    authority=AUTHORITY,
    client_credential=GRAPH_APP_CLIENT_SECRET
)
token_result = msal_app.acquire_token_for_client(SCOPES)
if "access_token" not in token_result:
    raise RuntimeError(f"Failed to get token: {token_result}")
access_token = token_result["access_token"]
print("✅ Token acquired.")



## 3) Resolve siteId and prepare headers

import requests

headers = {
    "Authorization": f"Bearer {access_token}",
    "Accept": "application/json"
}

print(f"🔎 Resolving siteId for https://{SP_TENANT_HOST}/sites/{SP_SITE_PATH} …")
site_url = f"https://graph.microsoft.com/v1.0/sites/{SP_TENANT_HOST}:/sites/{SP_SITE_PATH}"
site_res = requests.get(site_url, headers=headers)
site_res.raise_for_status()
site_id = site_res.json()["id"]
print("✅ siteId:", site_id)



## 4) Grant site access (Sites.Selected) — grants App B; can also grant App A if needed
# This cell:
# 1) Checks existing app grants on the site
# 2) Grants "write" to App B (TARGET_APP_ID) if missing
# 3) (Optional) Grants "write" to App A (caller) if GRANT_CALLER_APP_TOO=True

def ensure_app_write_grant(site_id: str, app_client_id: str, display_name: str):
    """Grant 'write' role to the given application on the site if not already granted."""
    grant_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/permissions"
    res = requests.get(grant_url, headers=headers)
    res.raise_for_status()

    already = False
    for perm in res.json().get("value", []):
        for g in perm.get("grantedToIdentitiesV2", []):
            if g.get("application", {}).get("id") == app_client_id:
                print(f"ℹ️ App {app_client_id} already granted with roles: {perm.get('roles')}")
                already = True

    if not already:
        print(f"🛂 Granting 'write' to app {app_client_id} on this site …")
        payload = {
            "roles": ["write"],
            "grantedToIdentities": [
                {"application": {"id": app_client_id, "displayName": display_name}}
            ]
        }
        create_res = requests.post(grant_url, headers=headers, json=payload)
        if create_res.status_code != 201:
            raise RuntimeError(f"Grant failed for {app_client_id}: {create_res.status_code} {create_res.text}")
        print("✅ Grant created.")

# 4a) Grant App B (target app) as requested
ensure_app_write_grant(site_id, TARGET_APP_CLIENT_ID, TARGET_APP_DISPLAY)

# 4b) Optionally also grant the calling app (App A) to avoid 403s during file operations in this notebook
if GRANT_CALLER_APP_TOO and GRAPH_APP_CLIENT_ID != TARGET_APP_CLIENT_ID:
    ensure_app_write_grant(site_id, GRAPH_APP_CLIENT_ID, "Notebook-Caller-App")



## 5) Get the Documents drive and resolve your target folder
# SharePoint’s default library is displayed as “Shared Documents” in the URL, but Graph exposes it as Documents.
# The code below finds the documents library, then resolves SP_TARGET_FOLDER_PATH under its root.

# Find a document library drive (prefer exact name 'Documents', fallback to first documentLibrary)
drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives?$select=id,name,driveType"
drv_res = requests.get(drives_url, headers=headers)
drv_res.raise_for_status()
drives = drv_res.json().get("value", [])

documents_drive_id = None
# prefer 'Documents'
for d in drives:
    if d.get("driveType") == "documentLibrary" and d.get("name") == "Documents":
        documents_drive_id = d["id"]
        break
# fallback
if not documents_drive_id:
    for d in drives:
        if d.get("driveType") == "documentLibrary":
            documents_drive_id = d["id"]
            print(f"ℹ️ Using document library: {d.get('name')}")
            break

if not documents_drive_id:
    raise RuntimeError("No SharePoint document library drive found on this site.")

print("✅ Documents driveId:", documents_drive_id)

# Resolve the target folder item
from urllib.parse import quote
encoded_folder_path = quote(SP_TARGET_FOLDER_PATH.strip("/"))
folder_probe_url = f"https://graph.microsoft.com/v1.0/drives/{documents_drive_id}/root:/{encoded_folder_path}"
folder_probe = requests.get(folder_probe_url, headers=headers)
if folder_probe.status_code != 200:
    raise RuntimeError(
        f"Target folder not found at 'Documents/{SP_TARGET_FOLDER_PATH}'. "
        f"Create it or correct SP_TARGET_FOLDER_PATH. Details: {folder_probe.status_code} {folder_probe.text}"
    )
target_folder_id = folder_probe.json()["id"]
print(f"✅ Target folder resolved: Documents/{SP_TARGET_FOLDER_PATH}")
print("   targetFolderId:", target_folder_id)



## 6) Recursively enumerate all files under the target folder
# This walks subfolders and collects every file item (id, downloadUrl, relative path).

def list_children_paged(list_url: str):
    """Yield children arrays across @odata.nextLink pages."""
    while list_url:
        r = requests.get(list_url, headers=headers)
        r.raise_for_status()
        body = r.json()
        yield body.get("value", [])
        list_url = body.get("@odata.nextLink")

def collect_files_recursive(drive_id: str, folder_id: str, base_rel_path: str = ""):
    """
    Depth-first traversal of a folder.
    Returns a list of dicts: {id, name, rel_path, downloadUrl}
    rel_path is the path under SP_TARGET_FOLDER_PATH (for mirroring on Lakehouse).
    """
    results = []
    children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/children"
    for page in list_children_paged(children_url):
        for it in page:
            name = it.get("name", "")
            if "folder" in it:
                # Recurse into subfolder
                sub_id = it["id"]
                sub_rel = f"{base_rel_path}/{name}" if base_rel_path else name
                results.extend(collect_files_recursive(drive_id, sub_id, sub_rel))
            elif "file" in it:
                results.append({
                    "id": it["id"],
                    "name": name,
                    "rel_path": base_rel_path,  # may be "" at top-level
                    "downloadUrl": it.get("@microsoft.graph.downloadUrl")
                })
    return results

print(f"📂 Scanning recursively under Documents/{SP_TARGET_FOLDER_PATH} …")
all_files = collect_files_recursive(documents_drive_id, target_folder_id, "")
print(f"✅ Found {len(all_files)} file(s) in Documents/{SP_TARGET_FOLDER_PATH} (recursive).")
for preview in all_files[:10]:
    rel = f"{SP_TARGET_FOLDER_PATH}/{preview['rel_path']}/{preview['name']}".replace("//","/")
    print(" •", rel)



## 7) Copy each file to Lakehouse Files (preserve subfolder structure)
# Use the base64 strategy mssparkutils.fs.put() writes text.
# Each print includes a comment-style mapping line showing the exact SharePoint → Lakehouse paths.

import os, base64
from notebookutils import mssparkutils

def lakehouse_dest_path(site_folder: str, rel_path: str, filename: str) -> str:
    # Mirror Documents/<SP_TARGET_FOLDER_PATH>/<rel_path>/<filename> under Files/<site>/<SP_TARGET_FOLDER_PATH>/<rel_path>/
    pieces = [LAKEHOUSE_FILES_ROOT, site_folder]
    if SP_TARGET_FOLDER_PATH:
        pieces.append(SP_TARGET_FOLDER_PATH.strip("/"))
    if rel_path:
        pieces.append(rel_path.strip("/"))
    pieces.append(filename)
    # Join with "/" to form OneLake-style path
    return "/".join(pieces).replace("//", "/")

def ensure_parent_dirs(full_path: str):
    # Make sure parent directories exist in Lakehouse Files
    parent = "/".join(full_path.split("/")[:-1])
    if parent and not mssparkutils.fs.exists(parent):
        mssparkutils.fs.mkdirs(parent)

def put_base64(path_in_lakehouse: str, raw_bytes: bytes):
    # Encode to base64 because fs.put writes text
    b64_text = base64.b64encode(raw_bytes).decode("utf-8")
    mssparkutils.fs.put(path_in_lakehouse, b64_text, overwrite=True)

copied = 0
for f in all_files:
    if not f.get("downloadUrl"):
        print(f"⚠️ Skipping (no downloadUrl): {f['name']}")
        continue

    # Compose readable SharePoint relative path for logging
    sp_rel = f"Documents/{SP_TARGET_FOLDER_PATH}/{f['rel_path']}/{f['name']}".replace("//","/")

    # Download file content
    dl = requests.get(f["downloadUrl"])
    if dl.status_code != 200:
        print(f"❌ Download failed: {sp_rel} (HTTP {dl.status_code})")
        continue

    # Compute Lakehouse path mirroring the SharePoint structure
    dest_path = lakehouse_dest_path(LAKEHOUSE_SITE_FOLDER, f["rel_path"], f["name"])
    ensure_parent_dirs(dest_path)

    # “Comment” line showing the exact mapping
    print(f"# COPY: SP '{sp_rel}'  ->  Lakehouse '{dest_path}'")
    put_base64(dest_path, dl.content)

    print(f"✅ Copied: {f['name']}")
    copied += 1

print(f"🎉 Completed. {copied} file(s) copied to Lakehouse.")
