// Diese Datei enthält die Logik für das Outlook Web-Panel.

const flowUrl = "https://script.google.com/macros/s/AKfycbxvwK283FepZ55cC3A5OFOs-7Os7PC4Bq9EvYlIT2pvD3u4gsnYChOBKXwdgo-z3Z0X/exec";

/**
 * Ruft den Flow auf und zeigt Rückmeldung im Web-Panel.
 */
async function callFlow(action) {
    const statusElement = document.getElementById("status-message");

    if (!Office.context.mailbox.item) {
        statusElement.innerText = "Fehler: Bitte wählen Sie eine E-Mail aus.";
        statusElement.style.color = "red";
        return;
    }

    statusElement.innerText = `Die Anforderung '${action}' wurde gestartet... (Dauer ca. 25 Sekunden)`;
    statusElement.style.color = "blue";

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

        statusElement.innerText = `Die Anforderung '${action}' wurde erfolgreich übermittelt.`;
        statusElement.style.color = "green";

    } catch (error) {
        console.error("Netzwerkfehler beim Flow-Call:", error);
        statusElement.innerText = "Ihre Anforderung wird gerade verarbeitet... (ca. 25 Sek.)";
        statusElement.style.color = "gray";
    }
}

// Klick-Events registrieren
Office.onReady(() => {
    document.getElementById("btnCreate").onclick = () => callFlow("create");
    document.getElementById("btnUpdate").onclick = () => callFlow("update");
    document.getElementById("btnClose").onclick = () => callFlow("close");
    console.log("Web-Panel für Ticket-Aktionen ist bereit.");
});
