import requests
import json

# -----------------------------
# CONFIGURATION
# -----------------------------
TENANT_ID = "YOUR_TENANT_ID"  # e.g. customertenant.onmicrosoft.com OR tenant GUID
CLIENT_ID = "YOUR_APPLICATION_ID"
CLIENT_SECRET = "YOUR_APPLICATION_SECRET"

GRAPH_BASE_URL = "https://graph.microsoft.com/beta"

# The single SharePoint site you want all users to follow
SPECIFIC_SITE_ID = "YOUR_SITE_ID"  # e.g. contoso.sharepoint.com,12345,abcdef...

# -----------------------------
# HELPER: GET ACCESS TOKEN
# -----------------------------
def get_access_token():
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    token_data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    resp = requests.post(token_url, data=token_data)
    resp.raise_for_status()
    return resp.json()["access_token"]

# -----------------------------
# HELPER: PAGINATED GET
# -----------------------------
def paginated_get(url, headers):
    results = []
    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return results

# -----------------------------
# HELPER: FOLLOW SITE
# -----------------------------
def follow_site(user_id, site_id, headers):
    follow_body = {"value": [{"id": site_id}]}
    follow_resp = requests.post(
        f"{GRAPH_BASE_URL}/users/{user_id}/followedSites/add",
        headers=headers,
        data=json.dumps(follow_body),
    )
    if follow_resp.status_code in [200, 204]:
        print(f"    ✅ User {user_id} now follows site {site_id}")
    else:
        print(f"    ❌ Failed for {user_id}: {follow_resp.text}")

# -----------------------------
# MAIN
# -----------------------------
def main():
    access_token = get_access_token()
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

    # 1. Get all users (with pagination)
    users_url = f"{GRAPH_BASE_URL}/users"
    users = paginated_get(users_url, headers)
    print(f"Found {len(users)} users")

    for user in users:
        user_id = user["id"]
        user_name = user.get("userPrincipalName", user_id)
        print(f"\nProcessing user: {user_name}")
        follow_site(user_id, SPECIFIC_SITE_ID, headers)

if __name__ == "__main__":
    main()
