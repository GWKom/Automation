// Diese Datei enthält die Logik für die Outlook-Buttons für Non-Premium Lizenz-Nutzer in Outlook-for-Web.

// Ihre Power Automate Flow URL:
const flowUrl = "https://defaultacee719ca1b74fd8919a21a0e90caf.d7.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/5823c2b964194a7893cfcc925fede8fe/triggers/manual/paths/invoke/?api-version=1&tenantId=tId&environmentName=Default-acee719c-a1b7-4fd8-919a-21a0e90cafd7";

/**
 * Ruft den Power Automate Flow auf.
 * @param {string} action - Die auszuführende Aktion ('create', 'update', 'close').
 */
async function callFlow(action) {
    if (!Office.context.mailbox.item) {
        console.error("Keine E-Mail kontextuell verfügbar.");
        return;
    }

    const payload = {
        messageId: Office.context.mailbox.item.itemId,
        action: action,
        userEmail: Office.context.mailbox.userProfile.emailAddress 
    };

    try {
        const response = await fetch(flowUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        console.log("Flow triggered:", action, response.status);
    } catch (error) {
        console.error("Fehler beim Aufruf des Flows:", error);
    }
}

Office.onReady(() => {
    console.log("Office.js bereit.");
});
