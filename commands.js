/* Send in 60 Seconds: Outlook function command. No external telemetry. */
(function () {
  "use strict";
  let inProgress = false;
  function notify(item, key, message, type) {
    try {
      item.notificationMessages.replaceAsync(key, {
        type: type === "error" ? "errorMessage" : "informationalMessage",
        message: message,
        persistent: true
      });
    } catch (_ignored) { /* Notifications are best effort. */ }
  }
  function sendIn60(event) {
    const finish = () => { inProgress = false; event.completed(); };
    if (inProgress) { finish(); return; }
    inProgress = true;
    const item = Office.context.mailbox.item;
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.15") ||
        !item || !item.delayDeliveryTime || typeof item.sendAsync !== "function") {
      if (item) notify(item, "delayError", "This Outlook version does not support the required Mailbox 1.15 send API.", "error");
      finish();
      return;
    }
    const scheduledAt = new Date(Date.now() + 60000);
    item.delayDeliveryTime.setAsync(scheduledAt, function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        notify(item, "delayError", "Unable to schedule: " + (result.error ? result.error.message : "unknown error"), "error");
        finish();
        return;
      }
      // Outlook may perform other send checks. Never claim that send succeeded here.
      item.sendAsync(function (_sendResult) {
        // Microsoft warns that post-send callback work may not run reliably.
        finish();
      });
    });
  }
  Office.onReady(function () {
    Office.actions.associate("sendIn60", sendIn60);
  });
}());
