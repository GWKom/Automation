const flowUrl = "https://script.google.com/macros/s/AKfycbxvwK283FepZ55cC3A5OFOs-7Os7PC4Bq9EvYlIT2pvD3u4gsnYChOBKXwdgo-z3Z0X/exec";

async function callFlow(action) {
    const statusElement = document.getElementById("status-message");

    if (!Office.context.mailbox.item) {
        statusElement.innerText = "Fehler: Keine E-Mail ausgewählt.";
        statusElement.style.color = "red";
        return;
    }

    statusElement.innerText = `Aktion '${action}' wird verarbeitet...`;
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

        const result = await response.json();

        if (response.ok && result.status === 200) {
            statusElement.innerText = `Die Anforderung '${action}' wurde erfolgreich gestartet...(Dauer ca. 25 Sekunden)`;
            statusElement.style.color = "green";
        } else {
            statusElement.innerText = `Hm, das dauert länger als gedacht. Fehler beim Flow: Status ${result.status}`;
            statusElement.style.color = "orange";
        }

    } catch (error) {
        console.error("Netzwerkfehler:", error);
        statusElement.innerText = "Oups, da hat was nicht funktioniert. Verbindungsfehler zum Flow.";
        statusElement.style.color = "red";
    }
}

Office.onReady(() => {
    document.getElementById("btnCreate").onclick = () => callFlow("create");
    document.getElementById("btnUpdate").onclick = () => callFlow("update");
    document.getElementById("btnClose").onclick = () => callFlow("close");
    console.log("Web-Panel Ticket-Aktionen bereit.");
});
