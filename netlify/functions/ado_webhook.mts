// --- File: netlify/functions/ado_webhook.mts ---

import type { Handler, HandlerEvent, HandlerContext } from "@netlify/functions";
import fetch from "node-fetch"; // Using node-fetch v3 for ESM compatibility
import { Buffer } from "node:buffer"; // For Base64 decoding

// --- Environment Variables & Constants ---
// Required: Set these Webhook URLs and ADO credentials in Netlify Env Vars
const GOOGLE_CHAT_WEBHOOK_HANS = process.env.GOOGLE_CHAT_WEBHOOK_HANS;
const GOOGLE_CHAT_WEBHOOK_ALEXIS = process.env.GOOGLE_CHAT_WEBHOOK_ALEXIS;
const GOOGLE_CHAT_WEBHOOK_JUSTIN = process.env.GOOGLE_CHAT_WEBHOOK_JUSTIN;
const GOOGLE_CHAT_WEBHOOK_EFFORT = process.env.GOOGLE_CHAT_WEBHOOK_EFFORT;
const ADO_WEBHOOK_USER = process.env.ADO_WEBHOOK_USER;
const ADO_WEBHOOK_PASS = process.env.ADO_WEBHOOK_PASS; // Use PAT for better security

// Hardcoded User IDs (as per previous code structure + new ID for Justin)
const CHAT_USER_ID_HANS = "110089480014983777747";
const CHAT_USER_ID_ALEXIS = "111538330948035296439";
const CHAT_USER_ID_JUSTIN = "114126982067491484128";

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

// Type for identifying the target user/webhook
type TargetUser = "hans" | "alexis" | "justin" | "effort" | null;

interface FormatResult {
  targetUser: TargetUser;
  payload: GoogleChatCardPayload | GoogleChatSimpleText | null;
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
    );
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
  // detailedMessageData not used reliably, use resource.fields.System.History
  const resource = payload.resource ?? {};
  const resourceFields = resource.fields ?? {};
  let project_name = resourceFields["System.TeamProject"] ?? null;
  if (!project_name) {
    project_name =
      payload.resourceContainers?.project?.name ?? "Unknown Project";
  }

  let cardOrTextPayload: GoogleChatCardPayload | GoogleChatSimpleText | null =
    null;
  let targetUser: TargetUser = null;

  try {
    if (eventType === "workitem.commented") {
      const wi_id = resource.id ?? "N/A";
      const wi_type = resourceFields["System.WorkItemType"] ?? "Work Item";
      const wi_title = resourceFields["System.Title"] ?? "N/A";
      const commenter =
        resourceFields["System.ChangedBy"]?.displayName ?? "Unknown User";

      // --- CORRECTED Comment Text Extraction ---
      let comment_text: string | null = null;
      const history_text = resourceFields["System.History"];
      if (
        history_text &&
        typeof history_text === "string" &&
        history_text.trim()
      ) {
        comment_text = history_text.trim();
        console.log(
          `DEBUG: Using System.History for comment text. Length: ${
            comment_text.length
          }. Excerpt: '${comment_text.substring(0, 70)}...'`
        );
      } else {
        console.warn(
          `DEBUG WARNING: System.History field missing or empty for WI #${wi_id}. Cannot process comment content.`
        );
        comment_text = null;
      }
      // --- End Corrected Extraction ---

      let work_item_link = "#";
      // Extract link (checking markdown first)
      const markdown_text = messageData.markdown;
      if (markdown_text) {
        const match = markdown_text.match(/\[.*?\]\((.*?)\)/);
        if (match?.[1]) work_item_link = match[1];
      }
      if (work_item_link === "#" && messageData.html) {
        // Fallback to HTML link
        const match = messageData.html.match(/<a href="(.*?)">/);
        if (match?.[1]) work_item_link = match[1].replace(/&amp;/g, "&");
      }
      if (work_item_link === "#") {
        // Fallback to resource link
        work_item_link = resource._links?.html?.href ?? resource.url ?? "#";
      }

      // --- Check for specific STRINGS or user tags ---
      if (comment_text) {
        // Proceed only if comment text was successfully extracted
        const effortString = "Please review the total effort"; // Removed period
        const tagHans = "@Hans Stechl2";
        const tagAlexis = "@Alexis Aguirre";
        const tagJustin = "@Justin Burniske";

        console.log(`Checking comment for WI #${wi_id}. Using extracted text.`);

        // --- PRIORITY 1: Check for Effort Review String ---
        if (comment_text.includes(effortString)) {
          targetUser = "effort";
          // Use the <users/ID> format for Hans, using the hardcoded constant
          const simpleTextMessage = `<users/${CHAT_USER_ID_HANS}> - ${effortString} - ${work_item_link}`;
          cardOrTextPayload = { text: simpleTextMessage };
          console.log(
            `'${effortString}' FOUND in comment for WI #${wi_id}. Target: ${targetUser}. Formatting simple text with User ID mention.`
          );

          // --- PRIORITY 2: Check for User Tags (only if effort string wasn't found) ---
        } else {
          // Determine target user based on tags
          if (comment_text.includes(tagHans)) {
            targetUser = "hans";
            console.log(
              `Tag '${tagHans}' FOUND in comment for WI #${wi_id}. Target: ${targetUser}`
            );
          } else if (comment_text.includes(tagAlexis)) {
            targetUser = "alexis";
            console.log(
              `Tag '${tagAlexis}' FOUND in comment for WI #${wi_id}. Target: ${targetUser}`
            );
          } else if (comment_text.includes(tagJustin)) {
            targetUser = "justin";
            console.log(
              `Tag '${tagJustin}' FOUND in comment for WI #${wi_id}. Target: ${targetUser}`
            );
          }

          // If one of the user tags was found, format the detailed card
          if (targetUser) {
            const card_header = {
              title: `New Comment on ${wi_type} #${wi_id}`, // Simplified title
              subtitle: `${wi_title} | By: ${commenter}`,
              imageUrl: "https://img.icons8.com/color/48/000000/comments.png",
              imageType: "CIRCLE" as const,
            };
            const widgets: any[] = [
              { textParagraph: { text: `<b>Project:</b> ${project_name}` } },
              {
                textParagraph: {
                  text: `<b>Comment:</b><br>${comment_text.replace(
                    /\n/g,
                    "<br>"
                  )}`,
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
            cardOrTextPayload = {
              cardsV2: [
                {
                  cardId: card_id,
                  card: { header: card_header, sections: [{ widgets }] },
                },
              ],
            };
          } else {
            // Neither effort string nor user tag found
            console.log(
              `Neither effort string nor relevant tags found for WI #${wi_id} in extracted text. Skipping notification.`
            );
          }
        }
      } else {
        // comment_text was null
        console.error(
          `Failed to extract valid comment text for WI #${wi_id}. Skipping checks.`
        );
        targetUser = null;
        cardOrTextPayload = null;
      }
    } else {
      console.log(`Event type '${eventType}' not handled. Skipping.`);
    }
  } catch (error) {
    console.error(
      `Error during formatAdoEventCard for event ${eventType}: ${error}`
    );
    targetUser = null;
    cardOrTextPayload = null;
  }

  // Return the result object
  return { targetUser, payload: cardOrTextPayload };
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
  const webhookUrlJustin = process.env.GOOGLE_CHAT_WEBHOOK_JUSTIN;
  const webhookUrlEffort = process.env.GOOGLE_CHAT_WEBHOOK_EFFORT;

  // Check if at least one target webhook is configured
  if (
    !webhookUrlHans &&
    !webhookUrlAlexis &&
    !webhookUrlJustin &&
    !webhookUrlEffort
  ) {
    console.error(
      "No Google Chat Webhook URLs (HANS, ALEXIS, JUSTIN, or EFFORT) are configured in environment variables."
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

    // Determine the actual webhook URL to use based on the target user/type
    let targetWebhookUrl: string | undefined | null = null;
    if (formatResult.targetUser === "hans") {
      targetWebhookUrl = webhookUrlHans;
      console.log("Determined target user: hans");
    } else if (formatResult.targetUser === "alexis") {
      targetWebhookUrl = webhookUrlAlexis;
      console.log("Determined target user: alexis");
    } else if (formatResult.targetUser === "justin") {
      targetWebhookUrl = webhookUrlJustin;
      console.log("Determined target user: justin");
    } else if (formatResult.targetUser === "effort") {
      targetWebhookUrl = webhookUrlEffort;
      console.log("Determined target type: effort");
    } else {
      console.log("No specific target user/type determined by formatter.");
    }

    // Send notification ONLY if a target was identified, a payload was formatted, AND a URL exists for that target
    if (targetWebhookUrl && formatResult.payload) {
      console.log(
        `DEBUG: Payload type determined: ${
          formatResult.targetUser === "effort" ? "SimpleText" : "CardV2"
        }`
      );
      console.log(
        "DEBUG: Sending Payload:",
        JSON.stringify(formatResult.payload, null, 2)
      ); // Added Debug Log

      console.log(
        `Attempting to send payload to target: ${
          formatResult.targetUser || "effort"
        }`
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
      // No notification needed
      console.log(
        `No notification required or target webhook URL missing for targetUser/type: ${formatResult.targetUser}`
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
    await sendToGoogleChat(
      webhookUrlHans ||
        webhookUrlAlexis ||
        webhookUrlJustin ||
        webhookUrlEffort,
      {
        // Send error to first available URL
        text: `DEBUG ERROR: Exception during handler processing: ${error}`,
      }
    );
    return {
      statusCode: 500,
      body: JSON.stringify({
        status: "error",
        message: "Internal Server Error processing webhook",
      }),
    };
  }
};
