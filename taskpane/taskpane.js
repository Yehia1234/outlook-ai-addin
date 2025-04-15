// Global state variables (example)
let currentItem = null;
let currentAction = null; // To know which action triggered the taskpane
let pendingActionData = null; // Store data needed for approval

// Use Office.onReady or Office.initialize
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        setStatus("Ready.");
        currentItem = Office.context.mailbox.item;

        // Determine which button was clicked (passed via manifest or other means)
        // Simple way: check button IDs if taskpane opens uniquely per button (might need refinement)
        // Better way (if needed): Pass state via OfficeRuntime.storage or query Office.context.mailbox.item properties
        const commandId = Office.context.roamingSettings.get('lastCommandId'); // Needs setting in button action if possible, or infer
        if (commandId) {
             handleCommand(commandId);
             Office.context.roamingSettings.remove('lastCommandId'); // Clean up
        } else {
            // Default view or guess based on context (read/compose)
            if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
                if (Office.context.mailbox.item.displayMode === Office.MailboxEnums.DisplayMode.Read) {
                     setStatus("Select an AI action (Reply, Summarize, Translate, Follow-Up).");
                } else {
                     setStatus("Select an AI action (Mood Scan, Translate, Settings).");
                }
            }
        }

        // Register button handlers within the taskpane HTML
        document.getElementById('save-api-key').onclick = saveApiKey;
        document.getElementById('approve-action').onclick = executeApprovedAction;
        document.getElementById('decline-action').onclick = declineAction;

        // --- Add specific action button handlers here ---
        // Example: document.getElementById('run-mood-scan-button').onclick = runMoodScan; (If you add buttons in taskpane)
        // Note: The ribbon buttons directly trigger the functions via handleCommand below

    } else {
        setStatus("Host not supported: " + info.host);
    }
});


// === Core Function to Handle Ribbon Button Clicks ===
// This function needs to be called based on which ribbon button opened the pane.
// You might need a mechanism to pass the button ID to the taskpane,
// e.g., using OfficeRuntime.storage or checking Office.context.ui.displayDialogAsync options if applicable.
// A simpler approach for this example is assuming the button click logic is handled here directly.
// We'll map the manifest button IDs to functions.

function handleCommand(commandId) {
    currentAction = commandId; // Store which action is active
    hideAllSections(); // Hide previous results
    hideApprovalArea();
    hideError();

    switch (commandId) {
        case 'msgReadReplyButton':
            generateReply();
            break;
        case 'msgReadSummarizeButton':
            summarizeEmail();
            break;
        case 'msgReadTranslateButton': // Translate in Read mode
        case 'msgComposeTranslateButton': // Translate in Compose mode
             translateEmail();
            break;
        case 'msgReadFollowUpButton':
            checkFollowUp();
            break;
        case 'msgComposeMoodButton':
            scanMood();
            break;
        case 'msgSettingsButton':
             showSettings();
            break;
        default:
            setStatus(`Unknown command: ${commandId}`);
    }
}

// === API Key Management ===

const API_KEY_STORAGE_KEY = "geminiApiKey";

async function getApiKey() {
    // **INSECURE for production** - Use a backend service instead.
    // For development/demo, retrieve from OfficeRuntime.storage
    try {
        const storedKey = await OfficeRuntime.storage.getItem(API_KEY_STORAGE_KEY);
        if (!storedKey) {
            showError("API Key not set. Please configure it in Settings.");
            showSettings(); // Guide user to settings
            return null;
        }
        return storedKey;
    } catch (error) {
         showError(`Error retrieving API key: ${error}`);
         return null;
    }
}

async function saveApiKey() {
    const apiKeyInput = document.getElementById('api-key-input');
    const newKey = apiKeyInput.value.trim();
    if (newKey) {
        try {
             await OfficeRuntime.storage.setItem(API_KEY_STORAGE_KEY, newKey);
             setStatus("API Key saved successfully.");
             // Optionally hide settings section after save
             // hideAllSections();
        } catch (error) {
             showError(`Error saving API key: ${error}`);
        }
    } else {
        showError("Please enter a valid API Key.");
    }
}

function showSettings() {
     hideAllSections();
     document.getElementById('settings-section').style.display = 'block';
     // Optionally load the current key into the input field
     getApiKey().then(key => {
         if(key) document.getElementById('api-key-input').value = key;
     });
     setStatus("Configure Settings");
}


// === Gemini API Call Function ===
async function callGeminiAPI(prompt) {
    const apiKey = await getApiKey();
    if (!apiKey) return null; // Stop if no key

    const API_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`;

    setStatus("Calling AI Assistant...");
    try {
        const response = await fetch(API_ENDPOINT, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                contents: [{ parts: [{ text: prompt }] }],
                // Add safety settings, generation config if needed
                 "safetySettings": [ // Example safety settings
                    { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
                    { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
                    { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
                    { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"}
                 ],
                 "generationConfig": { // Example generation config
                    "temperature": 0.7, // Controls randomness (0=deterministic, 1=max random)
                    "maxOutputTokens": 1024 // Limit response size
                 }
            }),
        });

        if (!response.ok) {
             const errorBody = await response.text();
             throw new Error(`API Error (${response.status}): ${errorBody}`);
        }

        const data = await response.json();

        // Basic check for response structure
        if (data.candidates && data.candidates.length > 0 && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts.length > 0) {
            setStatus("AI response received.");
            return data.candidates[0].content.parts[0].text;
        } else if (data.promptFeedback && data.promptFeedback.blockReason) {
             throw new Error(`Content blocked by API safety filters: ${data.promptFeedback.blockReason}`);
        } else {
             console.error("Unexpected API response format:", data);
            throw new Error("Unexpected response format from Gemini API.");
        }

    } catch (error) {
        console.error("Gemini API Call Error:", error);
        showError(`Failed to call AI: ${error.message}`);
        setStatus("Error occurred.");
        return null;
    }
}

// === Feature Implementations ===

// --- Mood Scanner ('Smart Mood Scanner for Emails') ---
async function scanMood() {
    if (!currentItem || Office.context.mailbox.item.displayMode !== Office.MailboxEnums.DisplayMode.Compose) {
        showError("Please open an email draft to scan its mood.");
        return;
    }
    setStatus("Scanning email mood...");
    currentItem.body.getAsync(Office.CoercionType.Text, async (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailText = result.value.trim();
            if (!emailText) {
                 setStatus("Draft is empty.");
                 return;
            }

            const prompt = `Analyze the tone of this email draft. Identify if it sounds potentially aggressive, passive, vague, or unclear. Provide a brief analysis (1-2 sentences) and, if improvements are needed, suggest 1-2 alternative rephrasing options for specific sentences or the overall message.

Email Draft:
---
${emailText}
---

Analysis: [Your analysis here]
Suggestions (if any): [Your suggestions here]`;

            const response = await callGeminiAPI(prompt);
            if (response) {
                // Simple parsing (improve with more robust regex or structure)
                const analysisMatch = response.match(/Analysis:\s*([\s\S]*?)(Suggestions:|---|$)/i);
                const suggestionsMatch = response.match(/Suggestions \(if any\):\s*([\s\S]*)/i);

                const analysisText = analysisMatch ? analysisMatch[1].trim() : "Could not parse analysis.";
                let suggestionText = suggestionsMatch ? suggestionsMatch[1].trim() : "";

                // Display results
                 const section = document.getElementById('mood-scanner-section');
                 document.getElementById('mood-analysis').innerText = analysisText;
                 document.getElementById('mood-suggestions').innerText = suggestionText || "No specific suggestions provided.";
                 section.style.display = 'block';
                 setStatus("Mood scan complete.");

                 // Prepare for approval IF suggestions exist
                 if (suggestionText && suggestionText !== "No specific suggestions provided.") {
                     pendingActionData = { type: 'applyMoodSuggestion', suggestion: suggestionText }; // Store suggestion for approval
                     showApprovalArea("Apply Suggested Changes?", `The AI suggests these changes:\n\n${suggestionText}\n\nApply them to your draft? (This will replace the current draft content)`);
                 }
            }
        } else {
            showError("Failed to get email body: " + result.error.message);
        }
    });
}

// --- Reply Generator ('Reply to the email') ---
async function generateReply() {
     if (!currentItem || Office.context.mailbox.item.displayMode !== Office.MailboxEnums.DisplayMode.Read) {
        showError("Please select an email to reply to.");
        return;
    }
     setStatus("Generating reply...");
     // Get necessary details (Subject, Body, From)
     const subject = currentItem.subject;
     const from = currentItem.from?.emailAddress || 'Unknown Sender'; // Basic sender info

     currentItem.body.getAsync(Office.CoercionType.Text, async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
             const originalBody = result.value.trim();
             if (!originalBody) {
                 setStatus("Original email body is empty.");
                 return;
             }

             const prompt = `Read the following email and draft a suitable, polite, and professional reply.

Original Email Subject: ${subject}
Original Email From: ${from}
Original Email Body:
---
${originalBody}
---

Draft Reply Body:`;

             const response = await callGeminiAPI(prompt);
             if (response) {
                 // Show generated reply for approval
                 pendingActionData = { type: 'insertReply', replyText: response };
                 showApprovalArea("Use this Draft Reply?", response);
                 // Also show in the dedicated section for viewing
                 const section = document.getElementById('reply-section');
                 document.getElementById('reply-text').value = response; // Show in textarea too
                 section.style.display = 'block';
                 setStatus("Reply draft generated.");

             }
          } else {
             showError("Failed to get original email body: " + result.error.message);
          }
     });
}

// --- Summarizer ('Email Summarizer') ---
async function summarizeEmail() {
    if (!currentItem || Office.context.mailbox.item.displayMode !== Office.MailboxEnums.DisplayMode.Read) {
        showError("Please select an email to summarize.");
        return;
    }
    setStatus("Summarizing email...");
    currentItem.body.getAsync(Office.CoercionType.Text, async (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailText = result.value.trim();
             if (!emailText) {
                 setStatus("Email body is empty.");
                 return;
            }

            const prompt = `Summarize the following email content. Identify the main points, any key decisions mentioned, and list any clear action items with who they might be assigned to (if mentioned).

Email Content:
---
${emailText}
---

Summary: [Your summary here]
Key Decisions: [List decisions here, or "None identified"]
Action Items: [List action items here, or "None identified"]`;

            const response = await callGeminiAPI(prompt);
            if (response) {
                // Simple parsing (improve with better structure/regex)
                const summaryMatch = response.match(/Summary:\s*([\s\S]*?)(Key Decisions:|Action Items:|$)/i);
                const decisionsMatch = response.match(/Key Decisions:\s*([\s\S]*?)(Action Items:|$)/i);
                const actionsMatch = response.match(/Action Items:\s*([\s\S]*)/i);

                document.getElementById('summary-content').innerText = summaryMatch ? summaryMatch[1].trim() : "Could not parse summary.";
                document.getElementById('summary-decisions').innerText = decisionsMatch ? decisionsMatch[1].trim() : "None identified.";
                document.getElementById('summary-actions').innerText = actionsMatch ? actionsMatch[1].trim() : "None identified.";

                document.getElementById('summary-section').style.display = 'block';
                setStatus("Email summarized.");
                // No approval needed for just viewing
            }
        } else {
            showError("Failed to get email body: " + result.error.message);
        }
    });
}

// --- Translator ('Smart Email Translator & Culture Advisor') ---
async function translateEmail() {
    setStatus("Translating email...");
    currentItem.body.getAsync(Office.CoercionType.Text, async (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailText = result.value.trim();
             if (!emailText) {
                 setStatus("Email body is empty.");
                 return;
            }

            const prompt = `Translate the following English text into Arabic. After the translation, provide 2-3 brief cultural tips for business communication relevant to the content or general professional interaction with someone from the Middle East (e.g., formality, directness, greetings).

Original English Text:
---
${emailText}
---

Arabic Translation: [Your translation here]

Cultural Tips:
1. [Tip 1]
2. [Tip 2]
3. [Tip 3, optional]`;

            const response = await callGeminiAPI(prompt);
            if (response) {
                // Simple parsing
                const translationMatch = response.match(/Arabic Translation:\s*([\s\S]*?)(Cultural Tips:|$)/i);
                const tipsMatch = response.match(/Cultural Tips:\s*([\s\S]*)/i);

                const translation = translationMatch ? translationMatch[1].trim() : "Could not parse translation.";
                const tips = tipsMatch ? tipsMatch[1].trim() : "Could not parse tips.";

                document.getElementById('translation-output').innerText = translation;
                document.getElementById('cultural-tips').innerText = tips;
                document.getElementById('translate-section').style.display = 'block';
                setStatus("Translation complete.");

                // Only offer to insert if in Compose mode
                 if (Office.context.mailbox.item.displayMode === Office.MailboxEnums.DisplayMode.Compose) {
                     pendingActionData = { type: 'insertTranslation', translation: translation };
                     showApprovalArea("Insert Translation?", `Insert the following Arabic translation into your draft? (This will replace the current draft content)\n\n${translation}`);
                 }
            }
        } else {
            showError("Failed to get email body: " + result.error.message);
        }
    });
}


// --- Follow-Up Bot ('Auto-Follow-Up Bot' - Manual Check Version) ---
async function checkFollowUp() {
     if (!currentItem || Office.context.mailbox.item.itemType !== Office.MailboxEnums.ItemType.Message || Office.context.mailbox.item.displayMode !== Office.MailboxEnums.DisplayMode.Read) {
        showError("Please select a *sent* email in the reading pane to check for follow-up.");
        // Ideally, check if it's actually in the Sent Items folder, but that's harder with basic Office.js
        return;
    }
     setStatus("Checking follow-up status (simulation)...");
     const section = document.getElementById('followup-section');
     const statusDiv = document.getElementById('followup-status');
     section.style.display = 'block';

     // **Simplification:** This cannot reliably check reply status client-side easily.
     // It would typically involve:
     // 1. Storing the ItemID and Sent Date when the user wants to track it.
     // 2. Using Microsoft Graph API to find replies to that specific email's ConversationID or InternetMessageId after the sent date.
     // For this example, we simulate asking the AI based on a *preset* number of days.

     const daysThreshold = 7; // Example: check if older than 7 days
     const sentDateTime = currentItem.dateTimeSent; // Get sent date/time

     if (!sentDateTime) {
         statusDiv.innerText = "Could not determine the sent date of this email.";
         return;
     }

     const now = new Date();
     const daysSinceSent = (now.getTime() - sentDateTime.getTime()) / (1000 * 60 * 60 * 24);

     if (daysSinceSent > daysThreshold) {
          // Ask AI to draft a reminder (optional, could just prompt user)
          const subject = currentItem.subject;
          const toRecipients = (await getRecipients(currentItem.to)).join(', '); // Helper needed

         const prompt = `The user sent an email with the subject "${subject}" to ${toRecipients} over ${Math.floor(daysSinceSent)} days ago and may not have received a response. Draft a short, polite follow-up reminder email.

Original Subject: ${subject}

Reminder Draft:`;

          statusDiv.innerText = `Email sent ${Math.floor(daysSinceSent)} days ago (threshold: ${daysThreshold} days). Checking for replies is not implemented in this version. Generating reminder draft...`;
         const reminderDraft = await callGeminiAPI(prompt);

         if(reminderDraft) {
              pendingActionData = { type: 'sendReminder', reminderText: reminderDraft, originalSubject: subject, originalRecipients: await getRecipients(currentItem.to) }; // Store necessary data
              showApprovalArea("Send Follow-Up Reminder?", `AI suggests this reminder draft:\n\n${reminderDraft}\n\nDo you want to open this in a new email draft?`);
         } else {
             statusDiv.innerText += "\nCould not generate reminder draft via AI.";
         }

     } else {
         statusDiv.innerText = `Email sent ${Math.floor(daysSinceSent)} days ago. It's within the ${daysThreshold}-day threshold. No follow-up action suggested currently.`;
     }
}

// Helper to get recipients as simple list
async function getRecipients(recipientField) {
     return new Promise((resolve) => {
         recipientField.getAsync((asyncResult) => {
             if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                 resolve(asyncResult.value.map(recipient => recipient.emailAddress));
             } else {
                 console.error("Failed to get recipients:", asyncResult.error);
                 resolve([]); // Return empty on error
             }
         });
     });
}


// === Approval Flow ===

function showApprovalArea(title, content) {
     document.getElementById('approval-title').innerText = title;
     document.getElementById('approval-content').innerText = content; // Use innerText for safety
     document.getElementById('approval-area').style.display = 'block';
}

function hideApprovalArea() {
    document.getElementById('approval-area').style.display = 'none';
    pendingActionData = null; // Clear pending action
}

function executeApprovedAction() {
    if (!pendingActionData) return;

    setStatus("Applying action...");
    hideApprovalArea();

    switch (pendingActionData.type) {
        case 'applyMoodSuggestion':
            // Replace current body with the suggestion (use with caution!)
            // For safety, maybe prepend/append or offer manual copy instead.
             currentItem.body.setAsync(pendingActionData.suggestion, { coercionType: Office.CoercionType.Text }, (result) => {
                 if (result.status === Office.AsyncResultStatus.Succeeded) {
                     setStatus("Suggestion applied to draft.");
                 } else {
                     showError("Failed to apply suggestion: " + result.error.message);
                 }
             });
            break;

        case 'insertReply':
             // Option 1: Display reply form pre-filled
             Office.context.mailbox.item.displayReplyForm(
                 {
                     'htmlBody': pendingActionData.replyText // Assume AI gives HTML or convert plain text
                 }
             );
             setStatus("Reply form opened.");
             // Option 2 (if wanting to insert into *current* open reply draft - less common):
             // currentItem.body.setAsync(pendingActionData.replyText, { coercionType: Office.CoercionType.Html }, ...);
            break;

         case 'insertTranslation':
              currentItem.body.setAsync(pendingActionData.translation, { coercionType: Office.CoercionType.Text }, (result) => {
                 if (result.status === Office.AsyncResultStatus.Succeeded) {
                     setStatus("Translation inserted into draft.");
                 } else {
                     showError("Failed to insert translation: " + result.error.message);
                 }
             });
             break;

         case 'sendReminder':
            // Open a new compose window pre-filled
            let reminderSubject = `Re: ${pendingActionData.originalSubject}`;
            // Basic check to avoid multiple "Re: Re:"
            if (!pendingActionData.originalSubject.toLowerCase().startsWith("re:")) {
                reminderSubject = `Following Up: ${pendingActionData.originalSubject}`;
            }

             Office.context.mailbox.displayNewMessageForm({
                 toRecipients: pendingActionData.originalRecipients, // Use stored recipients
                 subject: reminderSubject,
                 htmlBody: pendingActionData.reminderText
             });
              setStatus("Reminder draft opened.");
             break;

        // Add cases for other actions needing approval
    }
     pendingActionData = null; // Clear action data after execution attempt
}

function declineAction() {
    hideApprovalArea();
    setStatus("Action declined.");
}


// === UI Helpers ===
function setStatus(message) {
    document.getElementById('status').innerText = message;
    console.log("Status:", message); // Log for debugging
}

function showError(message) {
     document.getElementById('error-message').innerText = message;
     document.getElementById('error-area').style.display = 'block';
     console.error("Error:", message);
}

function hideError() {
    document.getElementById('error-area').style.display = 'none';
}

function hideAllSections() {
    const sections = document.querySelectorAll('#action-area > div');
    sections.forEach(section => section.style.display = 'none');
     hideApprovalArea(); // Also hide approval when switching sections
}