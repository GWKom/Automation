const flowUrl = "https://script.google.com/macros/s/AKfycbxvwK283FepZ55cC3A5OFOs-7Os7PC4Bq9EvYlIT2pvD3u4gsnYChOBKXwdgo-z3Z0X/exec";

function setupButtons() {
    const createBtn = document.getElementById("btnCreate");
    const updateBtn = document.getElementById("btnUpdate");
    const closeBtn = document.getElementById("btnClose");

    if (createBtn && updateBtn && closeBtn) {
        createBtn.onclick = () => callFlow("create");
        updateBtn.onclick = () => callFlow("update");
        closeBtn.onclick = () => callFlow("close");

        console.log("✅ Buttons erfolgreich registriert.");
    } else {
        console.warn("⚠️ Buttons noch nicht verfügbar – retry in 500ms...");
        setTimeout(setupButtons, 500);
    }
}

async function callFlow(action) {
    const statusElement = document.getElementById("status-message");

    if (!Office.context.mailbox.item) {
        statusElement.innerText = "Fehler: Keine E-Mail ausgewählt.";
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
        await fetch(flowUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            mode: 'no-cors',
            body: JSON.stringify(payload)
        });

        statusElement.innerText = `Die Anforderung '${action}' wurde erfolgreich übermittelt.`;
        statusElement.style.color = "green";

    } catch (error) {
        console.error("Netzwerkfehler beim Flow-Call:", error);
        statusElement.innerText = "Ihre Anforderung wird gerade verarbeitet... (ca. 25 Sek.).";
        statusElement.style.color = "gray";
    }
}

// Office.js Initialisierung
if (typeof Office !== "undefined") {
    Office.onReady(info => {
        if (info.host === Office.HostType.Outlook) {
            console.log("📬 Office.js ist bereit.");
            setupButtons();
        }
    });
} else {
    console.warn("⚠️ Office.js nicht gefunden – fallback auf DOMContentLoaded.");
    document.addEventListener("DOMContentLoaded", setupButtons);
}
