import requests
from datetime import datetime, timedelta
import msal

# Azure AD App details
TENANT_ID = "98b65c17-ad0e-4e4d-9e97-0c9aaea188e9"
CLIENT_ID = "ff76bdb4-f137-4b4c-a5d9-fbe5e1e5686e"
CLIENT_SECRET = "mAB8Q~rzRjMIHopeOj6hl6UnPJAtwb4M.-wQMcUJ"
SCOPE = ["https://graph.microsoft.com/.default"]
CHAT_ID = "19:64e648b48cda43e2845bcd48d3b5adfd@thread.v2"
USER_NAME = "Aakarshit Rathore"

# Power Automate trigger URL
LOGIC_APP_URL = "https://prod-142.westus.logic.azure.com:443/workflows/6e3d1344be0a4cffa901c10ef97fe2a7/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=8ZeJGWYzYn_lNfAVIBaJEU4i1yFPlqLxQHB9aBGTKcY"

def get_access_token():
    app_auth = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    token = app_auth.acquire_token_for_client(scopes=SCOPE)
    return token.get("access_token")

def send_auto_reply():
    payload = {
        "message": "Please be patient, Aakarshit is reviewing the question."
    }
    response = requests.post(LOGIC_APP_URL, json=payload)
    
    if response.status_code in [200, 202]:
        print("✅ Power Automate Flow triggered successfully.")
    else:
        print(f"❌ Failed to trigger Flow: {response.status_code}, {response.text}")

def check_chat():
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    chat_url = f"https://graph.microsoft.com/v1.0/chats/{CHAT_ID}/messages"
    response = requests.get(chat_url, headers=headers)

    if response.status_code != 200:
        print("❌ Error getting messages:", response.status_code, response.text)
        return

    messages = response.json().get("value", [])
    messages = sorted(messages, key=lambda x: x.get("lastModifiedDateTime", ""), reverse=True)

    now = datetime.utcnow()
    latest_question = None
    latest_question_time = None

    # Step 1: Find latest question not from Aakarshit
    for msg in messages:
        msg_from = msg.get("from")
        msg_user = msg_from.get("user") if msg_from else None
        sender = msg_user.get("displayName", "") if msg_user else ""
        content = msg.get("body", {}).get("content", "").replace("<p>", "").replace("</p>", "")
        time_sent_raw = msg.get("lastModifiedDateTime", "")
        if not time_sent_raw:
            continue

        try:
            time_sent = datetime.strptime(time_sent_raw, "%Y-%m-%dT%H:%M:%S.%fZ")
        except ValueError:
            time_sent = datetime.strptime(time_sent_raw, "%Y-%m-%dT%H:%M:%SZ")

        if "?" in content and sender != USER_NAME:
            latest_question = {
                "time": time_sent,
                "content": content
            }
            print(f"✅ Found latest question: '{content}' sent at {time_sent}")
            break  # Get only the most recent question

    if not latest_question:
        print("✅ OK — No unanswered questions.")
        return {"status": "OK"}

    # Step 2: Check if Aakarshit or the bot replied after the question
    user_replied = False
    bot_replied = False
    for msg in messages:
        msg_from = msg.get("from")
        msg_user = msg_from.get("user") if msg_from else None
        sender = msg_user.get("displayName", "") if msg_user else ""
        content = msg.get("body", {}).get("content", "").replace("<p>", "").replace("</p>", "")
        time_sent_raw = msg.get("lastModifiedDateTime", "")
        if not time_sent_raw:
            continue

        try:
            time_sent = datetime.strptime(time_sent_raw, "%Y-%m-%dT%H:%M:%S.%fZ")
        except ValueError:
            time_sent = datetime.strptime(time_sent_raw, "%Y-%m-%dT%H:%M:%SZ")

        # Check if Aakarshit replied after the question
        if sender == USER_NAME and time_sent > latest_question["time"]:
            user_replied = True
            print(f"✅ Aakarshit has replied: '{content}' sent at {time_sent}")
        
        # Check if the bot (Workflows) replied
        if msg_from.get("application") and msg_from["application"].get("displayName") == "Workflows" and time_sent > latest_question["time"]:
            bot_replied = True
            print(f"✅ Bot (Workflows) has replied: '{content}' sent at {time_sent}")

        # If either replied, break the loop early
        if user_replied or bot_replied:
            break

    time_since_question = now - latest_question["time"]
    print(f"⏰ Time since question: {time_since_question}")

    # If the bot or Aakarshit has replied
    if user_replied or bot_replied:
        print("✅ OK — Aakarshit or the bot has replied to the question.")
        return {"status": "OK"}

    # Check if the time passed since the question is under 1 minute
    elif time_since_question < timedelta(minutes=1):
        print("✅ OK — Waiting for the answer from Aakarshit or the bot.")
        return {"status": "OK — Waiting for the answer"}

    else:
        print(f"⚠️ No reply from {USER_NAME} or the bot for: '{latest_question['content']}'")
        send_auto_reply()
        return {"status": f"No reply from {USER_NAME} or the bot", "question": latest_question["content"]}

# Run the script
if __name__ == "__main__":
    check_chat()