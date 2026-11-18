SEND IN 60 SECONDS - New Outlook for Windows (prototype)

WHAT IT DOES
Adds a ribbon command to a new/reply compose window. Clicking it sets a
server-side deferred delivery for 60 seconds in the future, then requests Send.
In New Outlook, scheduled messages should be in Drafts until delivery.

IMPORTANT: NOT INSTALLABLE DIRECTLY FROM THIS ZIP.
A custom add-in manifest is NOT a bundled executable. Its referenced commands.html,
commands.js and PNG images MUST be hosted at a reachable HTTPS URL. The XML file
must then be updated with the actual HTTPS URL before installing.

SETUP (requires a hosting location such as GitHub Pages)
1. Upload commands.html, commands.js, and assets/ folder to a HTTPS web host.
2. Edit manifest.xml and replace every occurrence of
   https://YOUR-HTTPS-HOST.example/send-in-60/
   with the public HTTPS base folder URL (including its final slash).
3. Open https://aka.ms/olksideload, select My add-ins -> Custom Addins ->
   Add a custom add-in -> Add from File and select manifest.xml.
4. Open New Outlook, create a harmless test message to YOURSELF, and look in
   compose ribbon's Apps/more actions for "Send in 60 Seconds".
5. Test on your mailbox and confirm Drafts, Sent Items and actual delivery time.
   If company policies restrict self-installed add-ins or APIs, contact IT.

LIMITATIONS / SAFETY
- Needs Outlook Mailbox requirement set 1.15 (and 1.13 for delayed delivery).
- Requires read/write mailbox permissions for sendAsync; review these before install.
- The user initiates sending by clicking the command. There is no repeating timer.
- The sendAsync callback is not a guarantee of delivery: other add-ins and mailbox
  rules can block it. On failures check Drafts and Outlook notifications.
- Clicking the ordinary Send button is unaffected; it sends normally.
- Test only with yourself first; deployment/hosting are not performed here.
- Other Outlook clients and mailbox configurations may not work.

Microsoft documentation:
https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/delay-delivery
https://learn.microsoft.com/en-us/javascript/api/outlook/office.messagecompose
https://learn.microsoft.com/en-us/outlook/add-ins/add-in-commands-for-outlook
