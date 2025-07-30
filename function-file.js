Office.onReady(() => {
  // Office.js is ready
});

function reportPhishing(event) {
  const item = Office.context.mailbox.item;
  item.forwardAsync({
    toRecipients: ["servicedesk@komplex-it.dk"],
    htmlBody: "<p>This email was reported as phishing.</p>"
  }, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "phishingReport",
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: "Phishing email forwarded.",
          icon: "icon16",
          persistent: false
        }
      );
    } else {
      console.error("Forward failed:", asyncResult.error.message);
    }
    event.completed();
  });
}

window.reportPhishing = reportPhishing;
