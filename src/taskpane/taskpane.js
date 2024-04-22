/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}



Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
      // Your initialization code here
      document.getElementById("sendButton").onclick = checkRecipientDomain;
  }
});

function checkRecipientDomain() {
  Office.context.mailbox.item.to.getAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          var recipients = result.value;
          for (var i = 0; i < recipients.length; i++) {
              var email = recipients[i].emailAddress;
              if (!isValidDomain(email)) {
                  alert("You're trying to send an email to an invalid domain.");
                  return;
              }
          }
          // If all recipients are valid, proceed with sending the email
          Office.context.mailbox.item.send();
      } else {
          console.error("Error retrieving recipient email addresses.");
      }
  });
}

function isValidDomain(email) {
  var predefinedDomain = "nxci.ca"; // Change this to your predefined domain
  var domain = email.split("@")[1];
  return domain === predefinedDomain;
}
