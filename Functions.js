// Diese Datei enthält die Logik für die Outlook-Buttons.

// Globale Definition der Flow-URL für einfache Wartung.
const flowUrl = "https://defaultacee719ca1b74fd8919a21a0e90caf.d7.environment.api.powerplatform.com/powerautomate/automations/direct/workflows/5823c2b964194a7893cfcc925fede8fe/triggers/manual/paths/invoke/?api-version=1&tenantId=acee719c-a1b7-4fd8-919a-21a0e90cafd7&environmentName=Default-acee719c-a1b7-4fd8-919a-21a0e90cafd7&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PHsGWrYrjRFivkmE3u8mIAtlZqDNEPXP1FUGCHmsfYE";

/**
 * Hilfsfunktion zum Aufrufen des Power Automate Flows
 * und zur Anzeige einer Erfolgsmeldung in Outlook.
 * @param {string} action - Die auszuführende Aktion ('create', 'update', 'close').
 */
async function callFlow(action) {
    if (!Office.context.mailbox.item) {
        console.error("Kein E-Mail-Kontext verfügbar.");
        return;
    }

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

        // NEU: Erfolgs- oder Fehlermeldung für den Benutzer anzeigen
        if (response.status === 202) { // 202 Accepted ist der Erfolgscode von Power Automate
            Office.context.mailbox.item.notificationMessages.addAsync(action + "_success", {
                type: "informational",
                message: `Aktion '${action}' wurde erfolgreich ausgelöst.`,
                icon: "icon16.create", // Sie können hier auch spezifische Icons definieren
                persistent: false
            });
        } else {
            Office.context.mailbox.item.notificationMessages.addAsync("flow_error", {
                type: "errorMessage",
                message: `Fehler beim Auslösen des Flows. Status: ${response.status}`
            });
        }

    } catch (error) {
        console.error("Netzwerkfehler beim Flow-Call:", error);
        Office.context.mailbox.item.notificationMessages.addAsync("network_error", {
            type: "errorMessage",
            message: "Netzwerkfehler. Bitte prüfen Sie Ihre Verbindung."
        });
    }
}

// Action-Handler-Funktionen, klar definiert für bessere Lesbarkeit
function createTicket(event) {
  callFlow("create");
  event.completed(); // Wichtig: Signal an Outlook, dass die Aktion beendet ist.
}

function updateTicket(event) {
  callFlow("update");
  event.completed();
}

function closeTicket(event) {
  callFlow("close");
  event.completed();
}

// Office.onReady registriert die Funktionen, wenn Outlook bereit ist.
Office.onReady(() => {
    Office.actions.associate("createTicket", createTicket);
    Office.actions.associate("updateTicket", updateTicket);
    Office.actions.associate("closeTicket", closeTicket);
    console.log("Ticket-Aktionen wurden erfolgreich für Outlook registriert.");
});