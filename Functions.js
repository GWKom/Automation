// Diese Datei enthält die Logik für die Outlook-Buttons.

// Ihre Power Automate Flow URL
const flowUrl = "https://defaultacee719ca1b74fd8919a21a0e90caf.d7.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/5823c2b964194a7893cfcc925fede8fe/triggers/manual/paths/invoke/?api-version=1&tenantId=tId&environmentName=Default-acee719c-a1b7-4fd8-919a-21a0e90cafd7";

/**
 * Eine Hilfsfunktion, die den Power Automate Flow aufruft.
 * @param {string} action - Die auszuführende Aktion ('create', 'update', 'close').
 */
async function callFlow(action) {
    if (Office.context.mailbox.item === null) { return; }

    const payload = {
        messageId: Office.context.mailbox.item.itemId,
        action: action,
        mailbox: Office.context.mailbox.userProfile.emailAddress,
        internetmessageId: Office.context.mailbox.item.internetMessageId
    };

    try {
        await fetch(flowUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
    } catch (error) {
        console.error(error);
    }
}

// Diese Funktionsnamen müssen exakt mit denen in der Automatisierung_Button_Outlook_GWKOM.xml übereinstimmen
function createTicket(event) {
    callFlow("create");
    event.completed(); // Signal an Outlook, dass die Aktion beendet ist.
}

function updateTicket(event) {
    callFlow("update");
    event.completed();
}

function closeTicket(event) {
    callFlow("close");
    event.completed();
}

// Die Funktionen für Outlook registrieren
Office.actions.associate("createTicket", createTicket);
Office.actions.associate("updateTicket", updateTicket);
Office.actions.associate("closeTicket", closeTicket);