/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2026 ClearCode Inc.
*/
import { ConfigLoader } from "./config-loader.mjs";
import { MailDataCreator } from "./mail-data-creator.mjs";

Office.onReady(() => {});

// function createNewMail() {
//   try {
//     const currentItemId = Office.context.mailbox.item.itemId;
//     MailDataCreator.CreateReplyMailData();
//     Office.context.mailbox.displayNewMessageFormAsync({
//       toRecipients: Office.context.mailbox.item.to, // Copies the To line from current item
//       ccRecipients: ["sam@contoso.com"],
//       subject: "Outlook add-ins are cool!",
//       htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
//       attachments: [
//         {
//           name: Office.context.mailbox.item.subject,
//           type: Office.MailboxEnums.AttachmentType.Item,
//           itemId: currentItemId,
//         },
//       ],
//     });
//   } catch (e) {
//     console.log("createNewMail Failed:", e);
//   }
// }

async function onTypicalReplyButtonClicked(event) {
  const actionId = event.source.id; 
  console.log(actionId);
  const config = await ConfigLoader.loadConfigForCurrentLanguage(Office.context.displayLanguage);
  const originalMailData = {
    toRecipients: Office.context.mailbox.item.to,
    ccRecipients: Office.context.mailbox.item.cc,
    bccRecipients: Office.context.mailbox.item.bcc,
    sender: Office.context.mailbox.item.sender,
    body: Office.context.mailbox.item.body,
    subject: Office.context.mailbox.item.subject,
    id: Office.context.mailbox.item.itemId,
  };
  try {
    const waitComplete = false;
    for(const buttonConfig of config.ButtonConfigList) {
      if (actionId !== buttonConfig.Id) {
        continue;
      }
      const replyMailData = MailDataCreator.CreateReplyMailData({ config: config.ButtonConfigList[0], originalMailData });
      Office.context.mailbox.displayNewMessageFormAsync(replyMailData);
      // displayNewMessageFormAsync will be canceled if event.completed() is called
      // before finishing displayNewMessageFormAsync. The event will be completed 
      // automatically after displayNewMessageFormAsync is called.
      waitComplete = true;
      break;
    }
    if(!waitComplete) {
      event.completed();
    }
  } catch (e) {
    console.log("createNewMail Failed:", e);
  }
}
window.onTypicalReplyButtonClicked = onTypicalReplyButtonClicked;
Office.actions.associate("onTypicalReplyButtonClicked", onTypicalReplyButtonClicked);
