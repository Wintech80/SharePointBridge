"""First-run setup: sign in to Microsoft 365 and register the Entra ID app the bridge uses.

Replaces the Azure CLI approach: signs in with Microsoft's well-known Azure CLI public client
(the same one `az login` uses) and calls Microsoft Graph directly to create the app registration,
its service principal and the permission grant.
"""
import os
import webbrowser
from urllib.parse import urlparse

import msal
import requests

BOOTSTRAP_CLIENT = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"  # Azure CLI public client
GRAPH = "https://graph.microsoft.com/v1.0"
GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
APP_NAME = os.environ.get("SPB_APP_NAME", "SharePoint Bridge")
PERMISSIONS = {  # delegated Microsoft Graph permissions
    "Sites.Read.All": "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
    "Files.Read.All": "df85f4d6-205c-4ac5-a5ea-6bf408dba283",
    "Files.ReadWrite.All": "5c28f0bf-8a70-41f1-8ab2-9032436ddb65",
}


def find_tenant(*hints):
    """Work out the tenant id from hints such as a SharePoint URL or an email address. Returns None if unknown."""
    candidates = []
    for hint in hints:
        hint = (hint or "").strip()
        if not hint:
            continue
        if "@" in hint:
            candidates.append(hint.split("@", 1)[1])
        else:
            host = urlparse(hint if "//" in hint else "//" + hint).netloc or hint
            if host.endswith(".sharepoint.com"):
                prefix = host.split(".")[0].removesuffix("-my")
                candidates.append(f"{prefix}.onmicrosoft.com")
            else:
                candidates.append(host)
    for domain in candidates:
        try:
            r = requests.get(f"https://login.microsoftonline.com/{domain}/v2.0/.well-known/openid-configuration", timeout=10)
            if r.ok:
                return r.json()["issuer"].rstrip("/").split("/")[-2]
        except requests.RequestException:
            pass
    return None


def sign_in(tenant):
    """Device-code sign-in to the given tenant. Returns (access_token, tenant_id)."""
    app = msal.PublicClientApplication(BOOTSTRAP_CLIENT, authority=f"https://login.microsoftonline.com/{tenant}")
    flow = app.initiate_device_flow(scopes=["https://graph.microsoft.com/.default"])
    if "user_code" not in flow:
        raise RuntimeError(f"Could not start sign-in: {flow}")
    print("\n" + flow["message"])
    try:
        webbrowser.open(flow["verification_uri"])
    except Exception:
        pass
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Sign-in failed: {result.get('error_description', result)}")
    return result["access_token"], result["id_token_claims"]["tid"]


class _Graph:
    def __init__(self, token):
        self.h = {"Authorization": f"Bearer {token}"}

    def call(self, method, path, **kw):
        r = requests.request(method, GRAPH + path, headers=self.h, **kw)
        if not r.ok:
            try:
                msg = r.json()["error"]["message"]
            except Exception:
                msg = r.text
            raise RuntimeError(f"{method} {path} -> {r.status_code}: {msg}")
        return r.json() if r.content else {}


def register(tenant, read_only=True):
    """Create (or update) the app registration in `tenant`. Returns (client_id, tenant_id, admin_consent_granted)."""
    token, tenant = sign_in(tenant)
    g = _Graph(token)
    scopes = ["Sites.Read.All", "Files.Read.All" if read_only else "Files.ReadWrite.All"]
    access = [{"resourceAppId": GRAPH_APP_ID, "resourceAccess": [{"id": PERMISSIONS[s], "type": "Scope"} for s in scopes]}]

    print(f"Looking for existing '{APP_NAME}' app registration...")
    found = g.call("GET", "/applications", params={"$filter": f"displayName eq '{APP_NAME}'"})["value"]
    if found:
        app = found[0]
        print("Found it - updating its permissions.")
        g.call("PATCH", f"/applications/{app['id']}", json={"requiredResourceAccess": access})
    else:
        print("Creating app registration...")
        app = g.call("POST", "/applications", json={
            "displayName": APP_NAME,
            "signInAudience": "AzureADMyOrg",
            "isFallbackPublicClient": True,
            "publicClient": {"redirectUris": ["http://localhost"]},
            "requiredResourceAccess": access,
        })
    client_id = app["appId"]

    sps = g.call("GET", "/servicePrincipals", params={"$filter": f"appId eq '{client_id}'"})["value"]
    sp = sps[0] if sps else g.call("POST", "/servicePrincipals", json={"appId": client_id})

    consented = False
    try:
        graph_sp = g.call("GET", "/servicePrincipals", params={"$filter": f"appId eq '{GRAPH_APP_ID}'"})["value"][0]
        scope_text = " ".join(scopes)
        grants = g.call("GET", "/oauth2PermissionGrants",
                        params={"$filter": f"clientId eq '{sp['id']}' and resourceId eq '{graph_sp['id']}'"})["value"]
        if grants:
            g.call("PATCH", f"/oauth2PermissionGrants/{grants[0]['id']}", json={"scope": scope_text})
        else:
            g.call("POST", "/oauth2PermissionGrants", json={
                "clientId": sp["id"], "consentType": "AllPrincipals",
                "resourceId": graph_sp["id"], "scope": scope_text})
        consented = True
        print("Admin consent granted.")
    except RuntimeError as e:
        print(f"Could not grant admin consent automatically ({e}).")
        print(f"An admin can approve it at: https://login.microsoftonline.com/{tenant}/adminconsent?client_id={client_id}")
        print("Otherwise you will be asked to approve it yourself the first time you sign in.")
    return client_id, tenant, consented
