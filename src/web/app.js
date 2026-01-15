/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2026 ClearCode Inc.
*/
//import { ConfigLoader } from "./config-loader.mjs";

let locale;
let language;

Office.onReady(() => {
  language = Office.context.displayLanguage;
  document.documentElement.setAttribute("lang", language);
});

function createNewMail() {
  try {
    const currentItemId = Office.context.mailbox.item.itemId;
    Office.context.mailbox.displayNewMessageFormAsync({
      toRecipients: Office.context.mailbox.item.to, // Copies the To line from current item
      ccRecipients: ["sam@contoso.com"],
      subject: "Outlook add-ins are cool!",
      htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
      attachments: [
        {
          name: Office.context.mailbox.item.subject,
          type: Office.MailboxEnums.AttachmentType.Item,
          itemId: currentItemId,
        },
      ],
    });
  } catch (e) {
    console.log("createNewMail Failed:", e);
  }
}

async function onTypicalReplyButtonClicked() {
  createNewMail();
}
window.onTypicalReplyButtonClicked = onTypicalReplyButtonClicked;
Office.actions.associate("onTypicalReplyButtonClicked", onTypicalReplyButtonClicked);
