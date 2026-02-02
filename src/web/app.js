/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2026 ClearCode Inc.
*/
import { ConfigLoader } from "./config-loader.mjs";
import { ReplayMailDataCreator } from "./mail-data-creator.mjs";
import { OfficeDataAccessHelper } from "./office-data-access-helper.mjs";

Office.onReady(() => {});

async function onTypicalReplyButtonClicked(event) {
  const actionId = event.source.id;
  console.debug("actionId: " + actionId);
  console.debug("conversationId: " + Office.context.mailbox.item.conversationId);
  const originalMailData = {
    toRecipients: Office.context.mailbox.item.to.map((recipients) => recipients.emailAddress),
    ccRecipients: Office.context.mailbox.item.cc.map((recipients) => recipients.emailAddress),
    bccRecipients: Office.context.mailbox.item.bcc.map((recipients) => recipients.emailAddress),
    sender: Office.context.mailbox.item.sender?.emailAddress,
    subject: Office.context.mailbox.item.subject,
    id: Office.context.mailbox.item.itemId,
  };
  try {
    const buttonConfig = await ConfigLoader.loadButtonConfig(
      Office.context.displayLanguage,
      actionId
    );
    if (!buttonConfig) {
      console.log("No button config find.");
      return event.completed();
    }
    if (!ReplayMailDataCreator.isAllRecipientsAllowed({ buttonConfig, originalMailData })) {
      console.log("Recipients contains some prohibited domains");
      return event.completed();
    }
    Office.context.roamingSettings.set(
      "conversationId",
      Office.context.mailbox.item.conversationId ?? ""
    );
    Office.context.roamingSettings.set("buttonconfig", JSON.stringify(buttonConfig));
    await OfficeDataAccessHelper.saveRoamingSettingsAsync();
    const attachments = ReplayMailDataCreator.getAttachments({ buttonConfig, originalMailData });
    const replyFormFunction = ReplayMailDataCreator.getReplyFormFunction(buttonConfig);
    replyFormFunction(
      {
        attachments,
      },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(`replyFormFunction failed with message ${asyncResult.error.message}`);
        }
        event.completed();
      }
    );
  } catch (e) {
    console.error("onTypicalReplyButtonClicked Failed:", e);
    event.completed();
  }
}
window.onTypicalReplyButtonClicked = onTypicalReplyButtonClicked;

async function onNewMessageComposeCreated(event) {
  const conversationId = Office.context.mailbox.item.conversationId;
  const buttonConfigString = Office.context.roamingSettings.get("buttonconfig")?.trim() ?? "";
  if (!buttonConfigString) {
    return event.completed();
  }
  const targetConversationId = Office.context.roamingSettings.get("conversationId")?.trim() ?? "";
  const buttonConfig = JSON.parse(buttonConfigString);
  if (conversationId !== targetConversationId) {
    return event.completed();
  }
  console.debug("conversation id matched.");
  Office.context.roamingSettings.remove("conversationId");
  Office.context.roamingSettings.remove("buttonconfig");
  await OfficeDataAccessHelper.saveRoamingSettingsAsync();

  const originalSubject = await OfficeDataAccessHelper.getSubjectAsync();
  const newSubject = ReplayMailDataCreator.createSubject({ buttonConfig, originalSubject });
  await OfficeDataAccessHelper.setSubjectAsync(newSubject);
  const recipients = ReplayMailDataCreator.getNewRecipients(buttonConfig);
  if (recipients.to) {
    await OfficeDataAccessHelper.setToAsync(recipients.to);
  }
  if (recipients.cc) {
    await OfficeDataAccessHelper.setCcAsync(recipients.cc);
  }
  if (recipients.bcc) {
    await OfficeDataAccessHelper.setBccAsync(recipients.bcc);
  }
  if (!buttonConfig.quoteType) {
    await OfficeDataAccessHelper.setBodyAsync("");
  }
  if (buttonConfig.body) {
    await OfficeDataAccessHelper.prependBodyAsync(`${buttonConfig.body} \n`);
  }
  event.completed();
}
window.onNewMessageComposeCreated = onNewMessageComposeCreated;

Office.actions.associate("onNewMessageComposeCreated", onNewMessageComposeCreated);
Office.actions.associate("onTypicalReplyButtonClicked", onTypicalReplyButtonClicked);
