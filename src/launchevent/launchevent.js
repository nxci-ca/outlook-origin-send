const customerDomain = "@nxci.ca";

function onMessageSendHandler(event) {
  let externalRecipients = [];
  
  // Get TO recipients
  Office.context.mailbox.item.to.getAsync(function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      var toRecipients = asyncResult.value;
      toRecipients.forEach(function(recipient) {
        if (!recipient.emailAddress.includes(customerDomain)) {
          externalRecipients.push(recipient.emailAddress);
        }
      });
    }

    // Get CC recipients
    Office.context.mailbox.item.cc.getAsync(function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        var ccRecipients = asyncResult.value;
        ccRecipients.forEach(function(recipient) {
          if (!recipient.emailAddress.includes(customerDomain)) {
            externalRecipients.push(recipient.emailAddress);
          }
        });
      }

      // Get BCC recipients
      Office.context.mailbox.item.bcc.getAsync(function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          var bccRecipients = asyncResult.value;
          bccRecipients.forEach(function(recipient) {
            if (!recipient.emailAddress.includes(customerDomain)) {
              externalRecipients.push(recipient.emailAddress);
            }
          });
        }

        // Check for external recipients
        if (externalRecipients.length > 0) {
          event.completed({
            allowEvent: false,
            errorMessage:
              "The mail includes some external recipients, are you sure you want to send it?\n\n" +
              externalRecipients.join("\n") +
              "\n\nClick Send to send the mail anyway.",
          });
        } else {
          event.completed({ allowEvent: true });
        }
      });
    });
  });
}

// To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
