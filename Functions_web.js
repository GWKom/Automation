// Diese Datei enthält die Logik für das Outlook Web-Panel.

// Globale Definition der Flow-URL
const flowUrl = "https://defaultacee719ca1b74fd8919a21a0e90caf.d7.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/5823c2b964194a7893cfcc925fede8fe/triggers/manual/paths/invoke/?api-version=1&tenantId=acee719c-a1b7-4fd8-919a-21a0e90cafd7&environmentName=Default-acee719c-a1b7-4fd8-919a-21a0e90cafd7&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PHsGWrYrjRFivkmE3u8mIAtlZqDNEPXP1FUGCHmsfYE";

/**
 * Hilfsfunktion zum Aufrufen des Power Automate Flows
 * und zur Anzeige einer Rückmeldung im Panel.
 */
async function callFlow(action) {
    const statusElement = document.getElementById("status-message");

    if (!Office.context.mailbox.item) {
        statusElement.innerText = "Fehler: Keine E-Mail ausgewählt.";
        return;
    }
    
    // Status-Meldung anzeigen, dass die Aktion läuft
    statusElement.innerText = `Aktion '${action}' wird ausgeführt...`;
    statusElement.style.color = "orange";

    const payload = {
        messageId: Office.context.mailbox.item.itemId,
        action: action,
        mailbox: Office.context.mailbox.userProfile.emailAddress,
        internetmessageId: Office.context.mailbox.item.internetMessageId
    };

    try {
        const response = await fetch(flowUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (response.status === 202) {
            statusElement.innerText = `Aktion '${action}' wurde erfolgreich ausgelöst!`;
            statusElement.style.color = "green";
        } else {
            statusElement.innerText = `Fehler beim Auslösen. Status: ${response.status}`;
            statusElement.style.color = "red";
        }

    } catch (error) {
        console.error("Netzwerkfehler beim Flow-Call:", error);
        statusElement.innerText = "Netzwerkfehler. Bitte Verbindung prüfen.";
        statusElement.style.color = "red";
    }
}

// Office.onReady stellt sicher, dass alles geladen ist, bevor wir
// die Klick-Events an unsere HTML-Buttons binden.
Office.onReady(() => {
    document.getElementById("btnCreate").onclick = () => callFlow("create");
    document.getElementById("btnUpdate").onclick = () => callFlow("update");
    document.getElementById("btnClose").onclick = () => callFlow("close");
    console.log("Web-Panel für Ticket-Aktionen ist bereit.");
});