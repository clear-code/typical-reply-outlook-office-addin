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
  console.debug("actionId: " + actionId);
  console.debug("conversationId: " + Office.context.mailbox.item.conversationId);
  const originalMailData = {
    toRecipients: Office.context.mailbox.item.to,
    ccRecipients: Office.context.mailbox.item.cc,
    bccRecipients: Office.context.mailbox.item.bcc,
    sender: Office.context.mailbox.item.sender,
    subject: Office.context.mailbox.item.subject,
    id: Office.context.mailbox.item.itemId,
  };
  try {
    const buttonConfig = await ConfigLoader.loadConfigForCurrentLanguageAndButtonId(
      Office.context.displayLanguage,
      actionId
    );
    if (!buttonConfig) {
      console.log("no button config find.");
      return event.completed();
    }
    const replyMailData = MailDataCreator.CreateDataOnForReplyForm({
      config: buttonConfig,
      originalMailData,
    });
    if (!replyMailData) {
      console.log("failed to create reply mail data.");
      return event.completed();
    }
    Office.context.roamingSettings.set(
      "conversationId",
      Office.context.mailbox.item.conversationId ?? ""
    );
    Office.context.roamingSettings.set("actionId", actionId);
    await OfficeDataAccessHelper.saveRoamingSettingsAsync();
    replyMailData.executeMethod({
      attachments: replyMailData.attachments,
      callback: () => {
        event.completed();
      },
    });
  } catch (e) {
    console.log("createNewMail Failed:", e);
    event.completed();
  }
}
window.onTypicalReplyButtonClicked = onTypicalReplyButtonClicked;

async function onNewMessageComposeCreated(event) {
  const conversationId = Office.context.mailbox.item.conversationId;
  const actionId = Office.context.roamingSettings.get("actionId")?.trim() ?? "";
  const targetConversationId = Office.context.roamingSettings.get("conversationId")?.trim() ?? "";
  console.debug("action id: " + actionId);
  console.debug("targetConversation id: " + targetConversationId);
  if (conversationId !== targetConversationId) {
    return event.completed();
  }
  const buttonConfig = await ConfigLoader.loadConfigForCurrentLanguageAndButtonId(
    Office.context.displayLanguage,
    actionId
  );
  Office.context.roamingSettings.remove("conversationId");
  Office.context.roamingSettings.remove("actionId");
  await OfficeDataAccessHelper.saveRoamingSettingsAsync();

  const currentSubject = await OfficeDataAccessHelper.getSubjectAsync();
  const data = MailDataCreator.CreateReplyMailData({ config: buttonConfig, currentSubject });
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
