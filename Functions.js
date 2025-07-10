const flowUrl = "https://script.google.com/macros/s/AKfycbxvwK283FepZ55cC3A5OFOs-7Os7PC4Bq9EvYlIT2pvD3u4gsnYChOBKXwdgo-z3Z0X/exec";

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
        const response = await fetch(flowUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const result = await response.json();
        const success = response.ok && result.status === 200;

        Office.context.mailbox.item.notificationMessages.addAsync(action + "_success", {
            type: "informational",
            message: `Die Anforderung '${action}' wurde erfolgreich gestartet...(Dauer ca. 25 Sekunden, Status ${result.status})`,
            icon: "icon16",
            persistent: false
        });

    } catch (error) {
        console.error("Netzwerkfehler beim Flow-Call:", error);
        Office.context.mailbox.item.notificationMessages.addAsync("network_error", {
            type: "errorMessage",
            message: "Oups, das dauert länger als gedacht. Bitte prüfen Sie Ihre Verbindung."
        });
    }

    event.completed();
}

function createTicket(event) { callFlow("create", event); }
function updateTicket(event) { callFlow("update", event); }
function closeTicket(event)  { callFlow("close", event); }

Office.onReady(() => {
    Office.actions.associate("createTicket", createTicket);
    Office.actions.associate("updateTicket", updateTicket);
    Office.actions.associate("closeTicket", closeTicket);
    console.log("Ticket-Aktionen erfolgreich registriert.");
});
