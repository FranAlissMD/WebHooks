// --- File: netlify/functions/ado_webhook.mts ---

import type { Handler, HandlerEvent, HandlerContext } from "@netlify/functions";
import fetch from 'node-fetch'; // Or remove if using native fetch in Node 18+
import { Buffer } from 'node:buffer'; // For Base64 decoding

// --- Environment Variables ---
const GOOGLE_CHAT_WEBHOOK_URL = process.env.GOOGLE_CHAT_WEBHOOK_URL;
const ADO_WEBHOOK_USER = process.env.ADO_WEBHOOK_USER;
const ADO_WEBHOOK_PASS = process.env.ADO_WEBHOOK_PASS; // Expecting a PAT

// --- Types for ADO Payload (simplified, add more detail if needed) ---
interface AdoPayload {
  eventType?: string;
  message?: { markdown?: string; html?: string; text?: string };
  detailedMessage?: { markdown?: string; html?: string; text?: string };
  resource?: any; // Using 'any' for simplicity, define specific types for better safety
  resourceContainers?: { project?: { name?: string } };
}

interface GoogleChatCardPayload {
  cardsV2: Array<{
    cardId: string;
    card: {
      header: {
        title: string;
        subtitle?: string;
        imageUrl?: string;
        imageType?: 'CIRCLE' | 'SQUARE';
      };
      sections: Array<{
        header?: string;
        widgets: Array<any>; // Define specific widget types for safety
      }>;
    };
  }>;
}

interface GoogleChatSimpleText {
    text: string;
}

// --- Helper to Send Message to Google Chat ---
async function sendToGoogleChat(messagePayload: GoogleChatCardPayload | GoogleChatSimpleText): Promise<boolean> {
  if (!GOOGLE_CHAT_WEBHOOK_URL) {
    console.error("GOOGLE_CHAT_WEBHOOK_URL environment variable not set. Cannot send message.");
    return false;
  }

  const headers = { 'Content-Type': 'application/json; charset=UTF-8' };
  try {
    console.log("Attempting to send payload to Google Chat:", JSON.stringify(messagePayload)); // Log outgoing payload
    const response = await fetch(GOOGLE_CHAT_WEBHOOK_URL, {
      method: 'POST',
      headers: headers,
      body: JSON.stringify(messagePayload),
      // timeout isn't a standard fetch option, handle via AbortController if needed
    });

    if (!response.ok) {
      const errorBody = await response.text();
      console.error(`Google Chat send FAILED! Status: ${response.status}, Body: ${errorBody}`);
      return false;
    } else {
      console.log(`Google Chat send SUCCEEDED! Status: ${response.status}`);
      return true;
    }
  } catch (error) {
    console.error(`EXCEPTION during fetch to Google Chat: ${error}`);
    return false;
  }
}

// --- Helper to Format ADO Event Card ---
function formatAdoEventCard(payload: AdoPayload): GoogleChatCardPayload | null {
    const eventType = payload.eventType ?? 'unknown';
    const messageData = payload.message ?? {};
    const detailedMessageData = payload.detailedMessage ?? {};
    const resource = payload.resource ?? {};
    const resourceFields = resource.fields ?? {};
    let project_name = resourceFields['System.TeamProject'] ?? null;
     if (!project_name) {
          project_name = payload.resourceContainers?.project?.name ?? 'Unknown Project';
     }

    let card: GoogleChatCardPayload | null = null; // Default to null

    try { // Add try block for safety during parsing
        if (eventType === 'workitem.commented') {
            const wi_id = resource.id ?? 'N/A';
            const wi_type = resourceFields['System.WorkItemType'] ?? 'Work Item';
            const wi_title = resourceFields['System.Title'] ?? 'N/A';
            const commenter = resourceFields['System.ChangedBy']?.displayName ?? 'Unknown User';

            let comment_text = "No comment text found.";
            const detailed_text = detailedMessageData.text;
            if (detailed_text) {
                const parts = detailed_text.split('\r\n').map(part => part.trim()).filter(part => part);
                if (parts.length > 0) comment_text = parts[parts.length - 1];
            }
            if (comment_text === "No comment text found.") { // Fallback
                comment_text = resourceFields['System.History'] ?? 'Could not retrieve comment text.';
            }

            let work_item_link = "#";
            const markdown_text = messageData.markdown;
            if (markdown_text) {
                const match = markdown_text.match(/\[.*?\]\((.*?)\)/);
                if (match?.[1]) work_item_link = match[1];
            }
            if (work_item_link === "#" && messageData.html) { // Fallback 1: HTML
                const match = messageData.html.match(/<a href="(.*?)">/);
                if (match?.[1]) work_item_link = match[1].replace(/&amp;/g, '&');
            }
            if (work_item_link === "#") { // Fallback 2: Resource Link
                work_item_link = resource._links?.html?.href ?? resource.url ?? '#';
            }

            const tag_to_find = '@Francisco Aliss';
            console.log(`Checking comment for WI #${wi_id}. Text: '${comment_text}'`); // Keep debug log
            if (comment_text.includes(tag_to_find)) {
                console.log(`Tag '${tag_to_find}' FOUND for WI #${wi_id}. Formatting card.`);
                const card_header = {
                    title: `New Comment on ${wi_type} #${wi_id}`,
                    subtitle: `${wi_title} | By: ${commenter}`,
                    imageUrl: "https://img.icons8.com/color/48/000000/comments.png",
                    imageType: 'CIRCLE' as const
                };
                const widgets: any[] = [
                    { textParagraph: { text: `<b>Project:</b> ${project_name}` } },
                    { textParagraph: { text: `<b>Comment:</b><br>${comment_text.replace(/\n/g, '<br>')}` } } // Ensure newlines are rendered
                ];
                if (work_item_link !== '#') {
                    widgets.push({ buttonList: { buttons: [{ text: "View Work Item", onClick: { openLink: { url: work_item_link } } }] } });
                }
                const card_id = `comment-wi-${wi_id}-rev-${resource.rev ?? 'N/A'}`;
                card = { cardsV2: [{ cardId: card_id, card: { header: card_header, sections: [{ widgets }] } }] };
            } else {
                console.log(`Tag '${tag_to_find}' NOT FOUND for WI #${wi_id}. Skipping.`);
            }

        } else if (eventType === 'git.pullrequest.created') {
            console.log("Processing 'git.pullrequest.created' event.");
            const repo = resource.repository ?? {};
            const repo_name = repo.name ?? 'N/A';
            project_name = repo.project?.name ?? project_name; // Refine project name

            const pr_id = resource.pullRequestId ?? 'N/A';
            const pr_title = resource.title ?? 'N/A';
            const creator = resource.createdBy?.displayName ?? 'N/A';
            const source_branch = (resource.sourceRefName ?? 'N/A').replace('refs/heads/', '');
            const target_branch = (resource.targetRefName ?? 'N/A').replace('refs/heads/', '');
            let pr_url = resource._links?.web?.href ?? '#';
            if (pr_url === '#') pr_url = resource.url ?? '#';

            const card_header = {
                title: `Pull Request #${pr_id} Created`,
                subtitle: `Repo: ${repo_name} | Project: ${project_name}`,
                imageUrl: "https://img.icons8.com/fluent/48/000000/pull-request.png",
                imageType: 'CIRCLE' as const
            };
            const widgets: any[] = [
                { decoratedText: { topLabel: "Title", text: pr_title } },
                { decoratedText: { topLabel: "Created By", text: creator } },
                { decoratedText: { topLabel: "Branches", text: `${source_branch} → ${target_branch}` } }
            ];
            if (pr_url !== '#') {
                widgets.push({ buttonList: { buttons: [{ text: "View Pull Request", onClick: { openLink: { url: pr_url } } }] } });
            }
            const card_id = `pr-${pr_id}`;
            card = { cardsV2: [{ cardId: card_id, card: { header: card_header, sections: [{ widgets }] } }] };
            console.log(`Formatted card for PR #${pr_id}.`);

        } else {
            console.log(`Event type '${eventType}' not handled. Skipping.`);
        }
    } catch (error) {
         console.error(`Error during formatAdoEventCard for event ${eventType}: ${error}`);
         sendToGoogleChat({text: `DEBUG ERROR: Exception during card formatting for event ${eventType}: ${error}`}); // Send error to chat
         card = null; // Ensure null is returned on error
    }

    return card;
}


// --- Netlify Function Handler ---
export const handler: Handler = async (event: HandlerEvent, context: HandlerContext) => {
  console.log("Handler execution started.");

  if (event.httpMethod !== "POST") {
    console.warn(`Received non-POST request: ${event.httpMethod}`);
    return {
      statusCode: 405,
      body: JSON.stringify({ status: "error", message: "Method Not Allowed" }),
    };
  }

  // --- Security Validation: Basic Authentication ---
  // --- Security Validation: Basic Authentication ---
  let authPassed = false;
  // Try accessing header case-insensitively (check lowercase first, then uppercase)
  const authHeader = event.headers.authorization || event.headers.Authorization;

  if (ADO_WEBHOOK_USER && ADO_WEBHOOK_PASS) {
    // Check if header exists and starts with 'basic ' (case-insensitive check)
    if (authHeader && authHeader.toLowerCase().startsWith("basic ")) {
      try {
        console.log("Attempting Basic Auth decoding.");
        // Correctly split into 'Basic' and 'Credentials', limit to 2 parts
        const parts = authHeader.split(" ", 2);
        if (parts.length === 2) {
          const encodedCredentials = parts[1];
          // Ensure encodedCredentials is a non-empty string before decoding
          if (encodedCredentials && typeof encodedCredentials === "string") {
            const decodedCredentials = Buffer.from(
              encodedCredentials,
              "base64"
            ).toString("utf-8");
            // Split into username and password (only at the first colon)
            const credentialParts = decodedCredentials.split(":", 2);
            if (credentialParts.length === 2) {
              const username = credentialParts[0];
              const password = credentialParts[1];

              if (
                username === ADO_WEBHOOK_USER &&
                password === ADO_WEBHOOK_PASS
              ) {
                console.log("Basic Authentication successful.");
                authPassed = true;
              } else {
                console.warn(
                  "Basic Authentication failed: Credentials mismatch."
                );
              }
            } else {
              console.warn(
                "Basic Authentication failed: Decoded credentials format invalid (missing colon?)."
              );
            }
          } else {
            console.warn(
              "Basic Authentication failed: Encoded credentials part is empty or not a string."
            );
          }
        } else {
          console.warn(
            "Basic Authentication failed: Authorization header format invalid (missing space?)."
          );
        }
      } catch (error) {
        // Catch errors during decoding/splitting
        console.error(`Error decoding/parsing Basic Auth header: ${error}`);
      }
    } else {
      // Header missing or doesn't start with 'basic '
      console.warn(
        "Basic Authentication failed: Missing or invalid Authorization header format."
      );
    }
  } else {
    // Required environment variables are missing
    console.error(
      "Basic Authentication environment variables (ADO_WEBHOOK_USER, ADO_WEBHOOK_PASS) are not configured."
    );
    // Keep authPassed = false, access will be denied
  }

  if (!authPassed) {
    console.error("Webhook Authentication Failed.");
    return {
      statusCode: 401, // Unauthorized
      body: JSON.stringify({
        status: "error",
        message: "Authentication Required",
      }),
    };
  }
  // --- End Security Validation ---

  // --- Process Payload if Auth Passed ---
  try {
    // event.body is null if no body, use '{}' as default string
    const payloadString = event.body ?? "{}";
    const payload: AdoPayload = JSON.parse(payloadString);
    const eventTypeReceived = payload.eventType ?? "unknown";
    console.log(`Webhook received for event type: ${eventTypeReceived}`);

    if (!GOOGLE_CHAT_WEBHOOK_URL) {
      console.error(
        "Google Chat Webhook URL is not configured, cannot send notification."
      );
      return {
        statusCode: 500, // Internal Server Error - configuration issue
        body: JSON.stringify({
          status: "error",
          message: "Webhook processor configuration error",
        }),
      };
    }

    console.log(
      `Formatting message for Google Chat (if applicable) for ${eventTypeReceived}...`
    );
    const chatPayload = formatAdoEventCard(payload); // Might return null

    if (chatPayload) {
      console.log("Attempting to send formatted card to Google Chat...");
      const sendSuccess = await sendToGoogleChat(chatPayload); // Use await

      if (sendSuccess) {
        return {
          statusCode: 200,
          body: JSON.stringify({
            status: "success",
            message: "Webhook received and notification sent",
          }),
        };
      } else {
        return {
          statusCode: 500, // Internal Server Error communicating with Chat
          body: JSON.stringify({
            status: "error",
            message:
              "Webhook received but failed to send notification to Google Chat",
          }),
        };
      }
    } else {
      // No payload formatted (event not handled, or comment didn't contain the tag)
      console.log(
        "No notification required for this specific event/condition."
      );
      return {
        statusCode: 200,
        body: JSON.stringify({
          status: "success",
          message: "Webhook received, no notification required",
        }),
      };
    }
  } catch (error) {
    // Catch errors during JSON parsing or general processing
    console.error(
      `Error processing webhook payload or formatting message: ${error}`
    );
    // Maybe try sending a basic error message to chat if possible
    await sendToGoogleChat({
      text: `DEBUG ERROR: Exception during handler processing: ${error}`,
    });
    return {
      statusCode: 500, // Internal Server Error
      body: JSON.stringify({
        status: "error",
        message: "Internal Server Error processing webhook",
      }),
    };
  }
};