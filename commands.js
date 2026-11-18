/* Send in 60 Seconds: Outlook function command. No external telemetry. */
(function () {
  "use strict";
  let inProgress = false;

  // Error notifications only. Outlook throws if "persistent" is set on an
  // errorMessage, and rejects messages longer than 150 characters.
  function notifyError(item, key, message) {
    try {
      item.notificationMessages.replaceAsync(key, {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: String(message).slice(0, 150)
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
      if (item) notifyError(item, "delayError", "This Outlook version does not support the required Mailbox 1.15 send API.");
      finish();
      return;
    }
    const scheduledAt = new Date(Date.now() + 60000);
    item.delayDeliveryTime.setAsync(scheduledAt, function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        notifyError(item, "delayError", "Unable to schedule: " + (result.error ? result.error.message : "unknown error"));
        finish();
        return;
      }
      // Outlook may perform other send checks. Never claim that send succeeded here.
      item.sendAsync(function (_sendResult) {
        // Microsoft: only event.completed() is guaranteed to run in this callback.
        finish();
      });
    });
  }

  Office.onReady(function () {
    Office.actions.associate("sendIn60", sendIn60);
  });
}());