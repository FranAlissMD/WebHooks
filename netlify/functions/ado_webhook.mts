// --- File: netlify/functions/ado_webhook.mts ---

import type { Handler, HandlerEvent, HandlerContext } from "@netlify/functions";
import fetch from "node-fetch"; // Using node-fetch v3 for ESM compatibility
import { Buffer } from "node:buffer"; // For Base64 decoding

// --- Environment Variables ---
const GOOGLE_CHAT_WEBHOOK_HANS = process.env.GOOGLE_CHAT_WEBHOOK_HANS;
const GOOGLE_CHAT_WEBHOOK_ALEXIS = process.env.GOOGLE_CHAT_WEBHOOK_ALEXIS;
const CHAT_USER_ID_HANS = "110089480014983777747";
const CHAT_USER_ID_ALEXIS = "111538330948035296439";
const ADO_WEBHOOK_USER = process.env.ADO_WEBHOOK_USER;
const ADO_WEBHOOK_PASS = process.env.ADO_WEBHOOK_PASS;

// --- Types ---
interface AdoPayload {
  eventType?: string;
  message?: { markdown?: string; html?: string; text?: string };
  detailedMessage?: { markdown?: string; html?: string; text?: string };
  resource?: any;
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
        imageType?: "CIRCLE" | "SQUARE";
      };
      sections: Array<{
        header?: string;
        widgets: Array<any>;
      }>;
    };
  }>;
}

interface GoogleChatSimpleText {
  text: string;
}

interface FormatResult {
  targetUser: "hans" | "alexis" | null;
  payload: GoogleChatCardPayload | null;
}

// --- Helper to Send Message to Google Chat ---
async function sendToGoogleChat(
  targetWebhookUrl: string | undefined | null,
  messagePayload: GoogleChatCardPayload | GoogleChatSimpleText
): Promise<boolean> {
  if (!targetWebhookUrl) {
    console.error(
      "Target Google Chat Webhook URL was not provided or is invalid. Cannot send message."
    );
    return false;
  }

  const headers = { "Content-Type": "application/json; charset=UTF-8" };
  try {
    console.log(
      `Attempting to send payload to Google Chat URL: ${targetWebhookUrl.substring(
        0,
        50
      )}...`
    ); // Log only prefix
    // console.debug("Outgoing Payload:", JSON.stringify(messagePayload)); // Optional: Log full payload if needed for deep debug
    const response = await fetch(targetWebhookUrl, {
      method: "POST",
      headers: headers,
      body: JSON.stringify(messagePayload),
    });

    if (!response.ok) {
      const errorBody = await response.text();
      console.error(
        `Google Chat send FAILED! Status: ${response.status}, Body: ${errorBody}`
      );
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
function formatAdoEventCard(payload: AdoPayload): FormatResult {
  const eventType = payload.eventType ?? "unknown";
  const messageData = payload.message ?? {};
  const detailedMessageData = payload.detailedMessage ?? {};
  const resource = payload.resource ?? {};
  const resourceFields = resource.fields ?? {};
  let project_name = resourceFields["System.TeamProject"] ?? null;
  if (!project_name) {
    project_name =
      payload.resourceContainers?.project?.name ?? "Unknown Project";
  }

  let cardPayload: GoogleChatCardPayload | null = null;
  let targetUser: FormatResult["targetUser"] = null; // Default target is null

  try {
    if (eventType === "workitem.commented") {
      const wi_id = resource.id ?? "N/A";
      const wi_type = resourceFields["System.WorkItemType"] ?? "Work Item";
      const wi_title = resourceFields["System.Title"] ?? "N/A";
      const commenter =
        resourceFields["System.ChangedBy"]?.displayName ?? "Unknown User";

      let comment_text = "No comment text found.";
      const detailed_text = detailedMessageData.text;
      if (detailed_text) {
        const parts = detailed_text
          .split("\r\n")
          .map((part) => part.trim())
          .filter((part) => part);
        if (parts.length > 0) comment_text = parts[parts.length - 1];
      }
      if (comment_text === "No comment text found.") {
        comment_text =
          resourceFields["System.History"] ??
          "Could not retrieve comment text.";
      }

      let work_item_link = "#";
      const markdown_text = messageData.markdown;
      if (markdown_text) {
        const match = markdown_text.match(/\[.*?\]\((.*?)\)/);
        if (match?.[1]) work_item_link = match[1];
      }
      if (work_item_link === "#" && messageData.html) {
        const match = messageData.html.match(/<a href="(.*?)">/);
        if (match?.[1]) work_item_link = match[1].replace(/&amp;/g, "&");
      }
      if (work_item_link === "#") {
        work_item_link = resource._links?.html?.href ?? resource.url ?? "#";
      }

      // --- Check for specific user tags ---
      const tagHans = "@Hans Stechl2";
      const tagAlexis = "@Alexis Aguirre";
      let mentionUserId: string | null = null;

      console.log(
        `Checking comment for WI #${wi_id}. Text excerpt: '${comment_text.substring(
          0,
          50
        )}...'`
      ); // Log excerpt
      if (comment_text.includes(tagHans)) {
        targetUser = "hans";
        mentionUserId = CHAT_USER_ID_HANS;
        console.log(
          `Tag '${tagHans}' FOUND for WI #${wi_id}. Target: ${targetUser}`
        );
      } else if (comment_text.includes(tagAlexis)) {
        targetUser = "alexis";
        mentionUserId = CHAT_USER_ID_ALEXIS;
        console.log(
          `Tag '${tagAlexis}' FOUND for WI #${wi_id}. Target: ${targetUser}`
        );
      }

      // If one of the tags was found, format the card
      if (targetUser && mentionUserId) {
        const mentionTag = mentionUserId.startsWith("YOUR_")
          ? ""
          : `<users/${mentionUserId}> `; // Use mention only if ID is valid

        const card_header = {
          title: `${mentionTag}New Comment on ${wi_type} #${wi_id}`,
          subtitle: `${wi_title} | By: ${commenter}`,
          imageUrl: "https://img.icons8.com/color/48/000000/comments.png",
          imageType: "CIRCLE" as const,
        };
        const widgets: any[] = [
          { textParagraph: { text: `<b>Project:</b> ${project_name}` } },
          {
            textParagraph: {
              text: `<b>Comment:</b><br>${comment_text.replace(/\n/g, "<br>")}`,
            },
          },
        ];
        if (work_item_link !== "#") {
          widgets.push({
            buttonList: {
              buttons: [
                {
                  text: "View Work Item",
                  onClick: { openLink: { url: work_item_link } },
                },
              ],
            },
          });
        }
        const card_id = `comment-wi-${wi_id}-rev-${resource.rev ?? "N/A"}`;
        cardPayload = {
          cardsV2: [
            {
              cardId: card_id,
              card: { header: card_header, sections: [{ widgets }] },
            },
          ],
        };
      } else {
        console.log(
          `Neither relevant tag found for WI #${wi_id}. Skipping notification.`
        );
        // targetUser and cardPayload remain null
      }
      // Removed the 'git.pullrequest.created' block
      // } else if (eventType === 'git.pullrequest.created') { ... }
    } else {
      console.log(`Event type '${eventType}' not handled. Skipping.`);
      // targetUser and cardPayload remain null
    }
  } catch (error) {
    console.error(
      `Error during formatAdoEventCard for event ${eventType}: ${error}`
    );
    // Maybe send error to a default/admin webhook if available?
    // sendToGoogleChat(ADMIN_WEBHOOK_URL_OR_DEFAULT, {text: `...`});
    targetUser = null;
    cardPayload = null; // Ensure null is returned on error
  }

  // Return the result object
  return { targetUser, payload: cardPayload };
}

// --- Netlify Function Handler ---
export const handler: Handler = async (
  event: HandlerEvent,
  context: HandlerContext
) => {
  console.log("Handler execution started.");

  if (event.httpMethod !== "POST") {
    console.warn(`Received non-POST request: ${event.httpMethod}`);
    return {
      statusCode: 405,
      body: JSON.stringify({ status: "error", message: "Method Not Allowed" }),
    };
  }

  // --- Security Validation: Basic Authentication ---
  let authPassed = false;
  const authHeader = event.headers.authorization || event.headers.Authorization;

  if (ADO_WEBHOOK_USER && ADO_WEBHOOK_PASS) {
    if (authHeader && authHeader.toLowerCase().startsWith("basic ")) {
      try {
        console.log("Attempting Basic Auth decoding.");
        const parts = authHeader.split(" ", 2);
        if (parts.length === 2) {
          const encodedCredentials = parts[1];
          if (encodedCredentials && typeof encodedCredentials === "string") {
            const decodedCredentials = Buffer.from(
              encodedCredentials,
              "base64"
            ).toString("utf-8");
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
        console.error(`Error decoding/parsing Basic Auth header: ${error}`);
      }
    } else {
      console.warn(
        "Basic Authentication failed: Missing or invalid Authorization header format."
      );
    }
  } else {
    console.error(
      "Basic Authentication environment variables (ADO_WEBHOOK_USER, ADO_WEBHOOK_PASS) are not configured."
    );
  }

  if (!authPassed) {
    console.error("Webhook Authentication Failed.");
    return {
      statusCode: 401,
      body: JSON.stringify({
        status: "error",
        message: "Authentication Required",
      }),
    };
  }
  // --- End Security Validation ---

  // Fetch specific webhook URLs needed later
  const webhookUrlHans = process.env.GOOGLE_CHAT_WEBHOOK_HANS;
  const webhookUrlAlexis = process.env.GOOGLE_CHAT_WEBHOOK_ALEXIS;

  if (!webhookUrlHans && !webhookUrlAlexis) {
    console.error(
      "No Google Chat Webhook URLs (HANS or ALEXIS) are configured in environment variables."
    );
    return {
      statusCode: 500,
      body: JSON.stringify({
        status: "error",
        message: "Webhook processor configuration error (No Target URL)",
      }),
    };
  }

  // --- Process Payload if Auth Passed ---
  try {
    const payloadString = event.body ?? "{}";
    const payload: AdoPayload = JSON.parse(payloadString);
    const eventTypeReceived = payload.eventType ?? "unknown";
    console.log(`Webhook received for event type: ${eventTypeReceived}`);

    console.log(
      `Formatting message for Google Chat (if applicable) for ${eventTypeReceived}...`
    );
    const formatResult = formatAdoEventCard(payload);

    let targetWebhookUrl: string | undefined | null = null;
    if (formatResult.targetUser === "hans") {
      targetWebhookUrl = webhookUrlHans;
      console.log("Determined target user: hans");
    } else if (formatResult.targetUser === "alexis") {
      targetWebhookUrl = webhookUrlAlexis;
      console.log("Determined target user: alexis");
    } else {
      console.log("No specific target user determined by formatter.");
    }

    if (targetWebhookUrl && formatResult.payload) {
      console.log(
        `Attempting to send formatted card to target: ${formatResult.targetUser}`
      );
      const sendSuccess = await sendToGoogleChat(
        targetWebhookUrl,
        formatResult.payload
      );

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
          statusCode: 500,
          body: JSON.stringify({
            status: "error",
            message:
              "Webhook received but failed to send notification to Google Chat",
          }),
        };
      }
    } else {
      console.log(
        `No notification required or target webhook URL missing for targetUser: ${formatResult.targetUser}`
      );
      return {
        statusCode: 200,
        body: JSON.stringify({
          status: "success",
          message: "Webhook received, no notification required/sent",
        }),
      };
    }
  } catch (error) {
    console.error(
      `Error processing webhook payload or formatting message: ${error}`
    );
    await sendToGoogleChat(webhookUrlHans || webhookUrlAlexis, {
      // Send error to first available URL
      text: `DEBUG ERROR: Exception during handler processing: ${error}`,
    });
    return {
      statusCode: 500,
      body: JSON.stringify({
        status: "error",
        message: "Internal Server Error processing webhook",
      }),
    };
  }
};