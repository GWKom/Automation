// Diese Datei enthält die Logik für die Outlook-Buttons.

// Globale Definition der Proxy-URL
const flowUrl = "https://script.google.com/macros/s/AKfycbxvwK283FepZ55cC3A5OFOs-7Os7PC4Bq9EvYlIT2pvD3u4gsnYChOBKXwdgo-z3Z0X/exec";

/**
 * Hilfsfunktion zum Aufrufen des Flows via Google Apps Script Proxy.
 * @param {string} action - 'create', 'update' oder 'close'.
 * @param {object} event - Das Event-Objekt vom Button-Klick.
 */
async function callFlow(action, event) {
    if (!Office.context.mailbox.item) {
        console.error("Kein E-Mail-Kontext verfügbar.");
        event.completed();
        return;
    }

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
            mode: 'no-cors',
            body: JSON.stringify(payload)
        });

        // Hinweis im Outlook-Fenster
        Office.context.mailbox.item.notificationMessages.addAsync(action + "_success", {
            type: "informational",
            message: `GW Kom Ticket-Automatisierung bearbeitet Ihre ${action}-Anforderung.`,
            icon: "icon16.create",
            persistent: false
        });

    } catch (error) {
        console.error("Netzwerkfehler beim Flow-Call:", error);
        Office.context.mailbox.item.notificationMessages.addAsync("network_error", {
            type: "errorMessage",
            message: "Netzwerkfehler. Bitte prüfen Sie Ihre Verbindung."
        });
    }

    event.completed(); // Immer zuletzt
}

// Klare Zuweisung für Outlook-Aktionen
function createTicket(event) {
    callFlow("create", event);
}
function updateTicket(event) {
    callFlow("update", event);
}
function closeTicket(event) {
    callFlow("close", event);
}

// Registriere die Aktionen
Office.onReady(() => {
    Office.actions.associate("createTicket", createTicket);
    Office.actions.associate("updateTicket", updateTicket);
    Office.actions.associate("closeTicket", closeTicket);
    console.log("Ticket-Aktionen wurden erfolgreich für Outlook registriert.");
});