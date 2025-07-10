// Diese Datei enthält die Logik für die Outlook-Buttons.

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
            body: JSON.stringify(payload)
        });

        // Visuelles Feedback im Outlook-Fenster
        Office.context.mailbox.item.notificationMessages.addAsync(action + "_started", {
            type: "informational",
            message: `Die Anforderung '${action}' wurde erfolgreich gestartet... (Dauer ca. 25 Sekunden)`,
            icon: "icon16",
            persistent: true
        });

    } catch (error) {
        console.error("Netzwerkfehler beim Flow-Call:", error);
        Office.context.mailbox.item.notificationMessages.addAsync("network_error", {
            type: "informational",
            message: "Ihre Anforderung wird gerade verarbeitet... (ca. 25 Sek.)",
            persistent: true
        });
    }

    event.completed(); // Wichtig!
}

// Handler-Funktionen
function createTicket(event) { callFlow("create", event); }
function updateTicket(event) { callFlow("update", event); }
function closeTicket(event) { callFlow("close", event); }

// Outlook-Registrierung
Office.onReady(() => {
    Office.actions.associate("createTicket", createTicket);
    Office.actions.associate("updateTicket", updateTicket);
    Office.actions.associate("closeTicket", closeTicket);
    console.log("Ticket-Aktionen wurden erfolgreich für Outlook registriert.");
});
