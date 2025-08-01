import os
import requests

CONFLUENCE_URL = "https://dhivakarnathk2993.atlassian.net/wiki"
PAGE_ID = "19726345"
USERNAME = os.getenv("CONFLUENCE_USER_EMAIL")
API_TOKEN = os.getenv("CONFLUENCE_API_KEY")

def update_confluence_page(hlsd_content):
    # Get current page version
    url = f"{CONFLUENCE_URL}/rest/api/content/{PAGE_ID}?expand=body.storage,version"
    response = requests.get(url, auth=(USERNAME, API_TOKEN))
    data = response.json()
    version = data["version"]["number"] + 1

    # Update page
    update_url = f"{CONFLUENCE_URL}/rest/api/content/{PAGE_ID}"
    headers = {"Content-Type": "application/json"}
    payload = {
        "id": PAGE_ID,
        "type": "page",
        "title": data["title"],
        "body": {
            "storage": {
                "value": f"<pre>{hlsd_content}</pre>",
                "representation": "storage"
            }
        },
        "version": {"number": version}
    }
    r = requests.put(update_url, json=payload, headers=headers, auth=(USERNAME, API_TOKEN))
    print(r.status_code, r.text)

if __name__ == "__main__":
    with open("hlsd.md", "r", encoding='utf-8') as f:
        hlsd = f.read()
    update_confluence_page(hlsd) 