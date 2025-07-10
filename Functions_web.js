// Diese Datei enthält die Logik für das Outlook Web-Panel.

// Proxy-URL via Google Apps Script
const flowUrl = "https://script.google.com/macros/s/AKfycbxvwK283FepZ55cC3A5OFOs-7Os7PC4Bq9EvYlIT2pvD3u4gsnYChOBKXwdgo-z3Z0X/exec";

/**
 * Aufruf des Proxys und Rückmeldung im Panel.
 */
async function callFlow(action) {
    const statusElement = document.getElementById("status-message");

    if (!Office.context.mailbox.item) {
        statusElement.innerText = "Fehler: Keine E-Mail ausgewählt.";
        return;
    }

    statusElement.innerText = `Aktion '${action}' wird ausgeführt...`;
    statusElement.style.color = "orange";

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

        // Auch ohne status-Auswertung Erfolg anzeigen
        statusElement.innerText = `Aktion '${action}' wurde übermittelt!`;
        statusElement.style.color = "green";

    } catch (error) {
        console.error("Netzwerkfehler beim Flow-Call:", error);
        statusElement.innerText = "Netzwerkfehler. Bitte Verbindung prüfen.";
        statusElement.style.color = "red";
    }
}

// Knöpfe registrieren
Office.onReady(() => {
    document.getElementById("btnCreate").onclick = () => callFlow("create");
    document.getElementById("btnUpdate").onclick = () => callFlow("update");
    document.getElementById("btnClose").onclick = () => callFlow("close");
    console.log("Web-Panel für Ticket-Aktionen ist bereit.");
});