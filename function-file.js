Office.onReady(() => {
    console.log("Support Add-In geladen");
});

function createSupportTicket(event) {
    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["support@kiener-wittlin.ch"],
    });
    event.completed();
}
