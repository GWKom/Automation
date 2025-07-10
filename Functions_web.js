// Diese Datei enthält die Logik für das Outlook Web-Panel.

// Globale Definition der Flow-URL
const flowUrl = "https://script.google.com/macros/s/AKfycbxvwK283FepZ55cC3A5OFOs-7Os7PC4Bq9EvYlIT2pvD3u4gsnYChOBKXwdgo-z3Z0X/exec";

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