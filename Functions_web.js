// Diese Datei enthält die Logik für das Outlook Web-Panel.

// Globale Definition der Proxy-URL
const flowUrl = "https://script.google.com/macros/s/AKfycbxvwK283FepZ55cC3A5OFOs-7Os7PC4Bq9EvYlIT2pvD3u4gsnYChOBKXwdgo-z3Z0X/exec";

/**
 * Dies ist die Haupt-Initialisierungsfunktion.
 * Office.js ruft diese Funktion automatisch auf, wenn die Bibliothek vollständig geladen ist.
 */
Office.initialize = function (reason) {
    // Stellt sicher, dass das HTML-Dokument auch bereit ist, bevor wir die Buttons verknüpfen.
    document.addEventListener("DOMContentLoaded", function() {
        document.getElementById("btnCreate").addEventListener("click", () => callFlow("create"));
        document.getElementById("btnUpdate").addEventListener("click", () => callFlow("update"));
        document.getElementById("btnClose").addEventListener("click", () => callFlow("close"));
        console.log("Web-Panel für Ticket-Aktionen ist bereit und die Buttons sind verknüpft.");
    });
};

/**
 * Ruft den Proxy auf und gibt Rückmeldung im Panel.
 */
async function callFlow(action) {
    const statusElement = document.getElementById("status-message");

    if (!Office.context.mailbox.item) {
        statusElement.innerText = "Fehler: Keine E-Mail ausgewählt.";
        statusElement.style.color = "red";
        return;
    }

    statusElement.innerText = `Die Anforderung '${action}' wurde erfolgreich gestartet...(Dauer ca. 25 Sekunden)`;
    statusElement.style.color = "blue";

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
            mode: 'no-cors',
            body: JSON.stringify(payload)
        });


        if (response.ok) {
            statusElement.innerText = `Die Anforderung '${action}' wurde erfolgreich übermittelt!`;
            statusElement.style.color = "green";
        } else {
            statusElement.innerText = `Fehler bei der Übermittlung. Status: ${response.status}`;
            statusElement.style.color = "red";
        }

    } catch (error) {
        console.error("Netzwerkfehler beim Flow-Call:", error);
        statusElement.innerText = "Ihre Anforderung wird gerade verarbeitet... (ca. 25 Sek.)";
        statusElement.style.color = "gray";
    }
}