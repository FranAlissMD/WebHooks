# --- File: netlify/functions/ado_webhook.py ---

import json
import requests
import re
import os
import logging
import base64

# --- Configuration via Environment Variables ---
# These MUST be set in Netlify's Environment Variables settings
GOOGLE_CHAT_WEBHOOK_URL = os.environ.get('GOOGLE_CHAT_WEBHOOK_URL')
ADO_WEBHOOK_USER = os.environ.get('ADO_WEBHOOK_USER')
ADO_WEBHOOK_PASS = os.environ.get('ADO_WEBHOOK_PASS') # Use PAT for better security

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Helper Function to Send Message to Google Chat ---
# (Same as before)
def send_to_google_chat(message_payload):
    """Sends a formatted payload to the configured Google Chat webhook."""
    if not GOOGLE_CHAT_WEBHOOK_URL:
        logging.error("GOOGLE_CHAT_WEBHOOK_URL environment variable not set. Cannot send message.")
        return False
    headers = {'Content-Type': 'application/json; charset=UTF-8'}
    try:
        response = requests.post(
            GOOGLE_CHAT_WEBHOOK_URL,
            json=message_payload,
            headers=headers,
            timeout=10
        )
        response.raise_for_status()
        logging.info("Successfully sent message to Google Chat.")
        return True
    except requests.exceptions.RequestException as e:
        logging.error(f"Failed to send message to Google Chat: {e}")
        return False
    except Exception as e:
        logging.exception("Unexpected error sending to Google Chat.")
        return False

# --- Helper Function to Format ADO Event for Google Chat (Card V2) ---
def format_ado_event_card(payload):
    """
    Formats specific Azure DevOps event data into a Google Chat Card V2 message.
    Returns None if the event type is not handled or conditions (e.g., tag) aren't met.
    """
    event_type = payload.get('eventType', 'unknown')
    message_data = payload.get('message', {})
    detailed_message_data = payload.get('detailedMessage', {})
    resource = payload.get('resource', {})
    resource_fields = resource.get('fields', {})
    project_name = resource_fields.get('System.TeamProject', None)
    if not project_name:
         project_name = payload.get('resourceContainers', {}).get('project', {}).get('name', 'Unknown Project')

    # Default return value if no specific handler processes the event
    chat_message = None
    card_header = {}
    widgets = []
    card_id = f"ado-event-{payload.get('id', 'default')}" # Default card ID

    # --- Handle workitem.commented Event (Conditional) ---
    if event_type == 'workitem.commented':
        wi_id = resource.get('id', 'N/A')
        wi_type = resource_fields.get('System.WorkItemType', 'Work Item')
        wi_title = resource_fields.get('System.Title', 'N/A')
        commenter = resource_fields.get('System.ChangedBy', {}).get('displayName', 'Unknown User')

        comment_text = "No comment text found."
        detailed_text = detailed_message_data.get('text')
        if detailed_text:
             parts = [part.strip() for part in detailed_text.split('\r\n') if part.strip()]
             if parts: comment_text = parts[-1]
        if comment_text == "No comment text found.":
            comment_text = resource_fields.get('System.History', 'Could not retrieve comment text.')

        work_item_link = "#"
        markdown_text = message_data.get('markdown')
        if markdown_text:
            match = re.search(r'\[.*?\]\((.*?)\)', markdown_text)
            if match: work_item_link = match.group(1)
        if work_item_link == "#":
             html_text = message_data.get('html')
             if html_text:
                  match = re.search(r'<a href="(.*?)">', html_text)
                  if match: work_item_link = match.group(1).replace('&amp;', '&')
        if work_item_link == "#":
             work_item_link = resource.get('_links', {}).get('html', {}).get('href', resource.get('url', '#'))

        # <<< CONDITION: Only format if the specific tag is in the comment >>>
        tag_to_find = '@Francisco Aliss'
        if tag_to_find in comment_text:
            logging.info(f"Tag '{tag_to_find}' found in comment for WI #{wi_id}. Formatting card.")
            card_header = {
                "title": f"New Comment on {wi_type} #{wi_id}",
                "subtitle": f"{wi_title} | By: {commenter}",
                "imageUrl": "https://img.icons8.com/color/48/000000/comments.png",
                "imageType": "CIRCLE"
            }
            widgets = [
                {"textParagraph": {"text": f"<b>Project:</b> {project_name}"}},
                {"textParagraph": {"text": f"<b>Comment:</b>\n{comment_text}"}}
            ]
            if work_item_link != '#':
                widgets.append({"buttonList": {"buttons": [{"text": "View Work Item", "onClick": {"openLink": {"url": work_item_link}}}]}})
            card_id = f"comment-wi-{wi_id}-rev-{resource.get('rev', 'N/A')}"
            # Construct the message for this specific case
            chat_message = {
                "cardsV2": [{"cardId": card_id, "card": {"header": card_header, "sections": [{"widgets": widgets}]}}]
            }
        else:
            logging.info(f"Tag '{tag_to_find}' not found in comment for WI #{wi_id}. Skipping notification.")
            # chat_message remains None

    # --- Handle git.pullrequest.created Event ---
    elif event_type == 'git.pullrequest.created':
        logging.info(f"Processing 'git.pullrequest.created' event.")
        repo = resource.get('repository', {})
        repo_name = repo.get('name', 'N/A')
        project_name = repo.get('project', {}).get('name', project_name)

        pr_id = resource.get('pullRequestId', 'N/A')
        pr_title = resource.get('title', 'N/A')
        creator = resource.get('createdBy', {}).get('displayName', 'N/A')
        source_branch = resource.get('sourceRefName', 'N/A').replace('refs/heads/', '')
        target_branch = resource.get('targetRefName', 'N/A').replace('refs/heads/', '')
        pr_url = resource.get('_links', {}).get('web', {}).get('href', '#')
        if pr_url == '#': pr_url = resource.get('url', '#')

        card_header = {
            "title": f"Pull Request #{pr_id} Created",
            "subtitle": f"Repo: {repo_name} | Project: {project_name}",
            "imageUrl": "https://img.icons8.com/fluent/48/000000/pull-request.png",
            "imageType": "CIRCLE"
        }
        widgets = [
            {"decoratedText": {"topLabel": "Title", "text": pr_title}},
            {"decoratedText": {"topLabel": "Created By", "text": creator}},
            {"decoratedText": {"topLabel": "Branches", "text": f"{source_branch} → {target_branch}"}}
        ]
        if pr_url != '#':
            widgets.append({"buttonList": {"buttons": [{"text": "View Pull Request", "onClick": {"openLink": {"url": pr_url}}}]}})
        card_id = f"pr-{pr_id}"
        # Construct the message for this specific case
        chat_message = {
            "cardsV2": [{"cardId": card_id, "card": {"header": card_header, "sections": [{"widgets": widgets}]}}]
        }

    # --- Event Type Not Handled ---
    else:
        logging.info(f"Event type '{event_type}' not configured for notification. Skipping.")
        # chat_message remains None

    return chat_message # Will be None if event type wasn't handled or condition not met

# --- Netlify Function Handler ---
# (Handler function remains the same - it already checks if chat_payload is None)
def handler(event, context):
    """
    Netlify Function handler for Azure DevOps webhooks.
    Validates using Basic Auth, processes payload, sends to Google Chat if applicable.
    """
    logging.info("Netlify function handler started.")

    # Check if it's a POST request
    if event.get('httpMethod') != 'POST':
        logging.warning(f"Received non-POST request: {event.get('httpMethod')}")
        return { 'statusCode': 405, 'body': json.dumps({'status': 'error', 'message': 'Method Not Allowed'}) }

    # --- Security Validation: Basic Authentication ---
    auth_passed = False
    auth_header = event.get('headers', {}).get('authorization')

    if ADO_WEBHOOK_USER and ADO_WEBHOOK_PASS:
        if auth_header and auth_header.lower().startswith('basic '):
            try:
                encoded_credentials = auth_header.split(' ', 1)[1]
                decoded_credentials = base64.b64decode(encoded_credentials).decode('utf-8')
                username, password = decoded_credentials.split(':', 1)

                if username == ADO_WEBHOOK_USER and password == ADO_WEBHOOK_PASS:
                    logging.info("Basic Authentication successful.")
                    auth_passed = True
                else:
                    logging.warning("Basic Authentication failed: Credentials mismatch.")
            except Exception as e:
                logging.error(f"Error decoding/parsing Basic Auth header: {e}")
        else:
            logging.warning("Basic Authentication failed: Missing or invalid Authorization header.")
    else:
        logging.error("Basic Authentication environment variables (ADO_WEBHOOK_USER, ADO_WEBHOOK_PASS) are not configured.")
        auth_passed = False

    if not auth_passed:
        logging.error("Webhook Authentication Failed.")
        return { 'statusCode': 401, 'body': json.dumps({'status': 'error', 'message': 'Authentication Required'}) }
    # --- End Security Validation ---

    # --- Process Payload if Auth Passed ---
    try:
        payload_string = event.get('body', '{}')
        payload = json.loads(payload_string)
        logging.info(f"Webhook received for event type: {payload.get('eventType', 'unknown')}")

        if not GOOGLE_CHAT_WEBHOOK_URL:
             logging.error("Google Chat Webhook URL is not configured, cannot send notification.")
             return { 'statusCode': 500, 'body': json.dumps({'status': 'error', 'message': 'Webhook processor configuration error'}) }

        logging.info("Formatting message for Google Chat (if applicable)...")
        chat_payload = format_ado_event_card(payload) # Might return None

        if chat_payload:
            logging.info("Sending message to Google Chat...")
            send_success = send_to_google_chat(chat_payload)
            if send_success:
                return { 'statusCode': 200, 'body': json.dumps({'status': 'success', 'message': 'Webhook received and notification sent'}) }
            else:
                return { 'statusCode': 500, 'body': json.dumps({'status': 'error', 'message': 'Webhook received but failed to send notification to Google Chat'}) }
        else:
            # No payload formatted (event not handled, or comment didn't contain the tag)
            logging.info("No notification required for this specific event/condition.")
            return { 'statusCode': 200, 'body': json.dumps({'status': 'success', 'message': 'Webhook received, no notification required'}) }

    except json.JSONDecodeError as e:
        logging.error(f"Failed to decode JSON payload: {e}")
        return { 'statusCode': 400, 'body': json.dumps({'status': 'error', 'message': 'Invalid JSON payload'}) }
    except Exception as e:
        logging.exception(f"Error processing webhook payload or formatting message: {e}")
        return { 'statusCode': 500, 'body': json.dumps({'status': 'error', 'message': 'Internal Server Error processing webhook'}) }