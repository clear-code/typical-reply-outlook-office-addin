/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2026 ClearCode Inc.
*/
import { ConfigLoader } from "./config-loader.mjs";
import { MailDataCreator } from "./mail-data-creator.mjs";
import { OfficeDataAccessHelper } from "./office-data-access-helper.mjs";

Office.onReady(() => {});

async function onTypicalReplyButtonClicked(event) {
  const actionId = event.source.id; 
  console.log(actionId);
  console.log("conversationId: " + Office.context.mailbox.item.conversationId);
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
  const abc = await OfficeDataAccessHelper.getAllMailData();
  // Office.context.mailbox.item.displayReplyAllFormAsync({
  //   htmlBody:
  // });
  console.log(Office.context.mailbox.item.itemId);
  try {
    const waitComplete = false;
    for(const buttonConfig of config.ButtonConfigList) {
      if (actionId !== buttonConfig.Id) {
        continue;
      }
      const replyMailData = MailDataCreator.CreateDataOnForReplyForm({ config: config.ButtonConfigList[0], originalMailData });
      if (!replyMailData) {
        return event.completed();
      }
      Office.context.roamingSettings.set(targetConversationId, Office.context.mailbox.item.conversationId ?? "");
      Office.context.roamingSettings.set("buttonConfig", JSON.stringify(config.ButtonConfigList[0]));
      Office.context.roamingSettings.saveAsync();
      replyMailData.executeMethod({attachments: replyMailData.attachments});
      //Office.context.mailbox.displayNewMessageFormAsync(replyMailData,);
      // displayNewMessageFormAsync will be canceled if event.completed() is called
      // before finishing displayNewMessageFormAsync. The event will be completed 
      // automatically after displayNewMessageFormAsync is called.
      waitComplete = true;
      break;
    }
    event.completed();
  } catch (e) {
    console.log("createNewMail Failed:", e);
  }
}
window.onTypicalReplyButtonClicked = onTypicalReplyButtonClicked;

async function onNewMessageComposeCreated(event) {
  const id =  Office.context.mailbox.item.conversationId;
  const targetConversationId = Office.context.roamingSettings.get(conversationId)?.trim() ?? "";
  if (id !== targetConversationId) {
    event.completed();
  }
  const buttonConfigJson = Office.context.roamingSettings.get("buttonConfig")?.trim() ?? "";
  Office.context.roamingSettings.remove(conversationId);
  Office.context.roamingSettings.remove("buttonConfig");
  Office.context.roamingSettings.saveAsync();
  const buttonConfig = JSON.parse(buttonConfigJson);
  const currentSubject = await OfficeDataAccessHelper.getSubjectAsync();
  const data = MailDataCreator.CreateReplyMailData({config: buttonConfig, currentSubject});
  if (data.newToRecipients) {
    await OfficeDataAccessHelper.setToAsync(data.newToRecipients);
  }
  await OfficeDataAccessHelper.setSubjectAsync(data.subject);
  let body = "";
  if (data.quoteType) {
    body = await OfficeDataAccessHelper.getBodyAsync();
  }
  if (data.bodyHtml) {
    body = data.bodyHtml + "\n\n" + body;
  }
  await OfficeDataAccessHelper.setBodyAsync(body);
  event.completed();
}
window.onNewMessageComposeCreated = onNewMessageComposeCreated;

Office.actions.associate("onNewMessageComposeCreated", onNewMessageComposeCreated);
Office.actions.associate("onTypicalReplyButtonClicked", onTypicalReplyButtonClicked);
