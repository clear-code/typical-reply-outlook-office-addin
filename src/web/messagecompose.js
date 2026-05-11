/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2026 ClearCode Inc.
*/
import { ReplayMailDataCreator } from "./mail-data-creator.mjs";
import { OfficeDataAccessHelper } from "./office-data-access-helper.mjs";

Office.onReady(() => {});

function plainTextToHtml(text) {
  if (!text) return "";

  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;")
    .replace(/\n/g, "<br>");
}

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
    const coercionType = await OfficeDataAccessHelper.getBodyTypeAsync();
    let prependBody = `${buttonConfig.body} \n`;
    if (coercionType === Office.CoercionType.Html) {
      prependBody = plainTextToHtml(prependBody);
    }
    await OfficeDataAccessHelper.prependBodyAsync(prependBody, coercionType);
    const body = await OfficeDataAccessHelper.getBodyAsync();
    // prependBodyAsync sometimes does not update the body immediately, so we set body again to make sure the body is updated.
    await OfficeDataAccessHelper.setBodyAsync(body, coercionType);
  }
  event.completed();
}
window.onNewMessageComposeCreated = onNewMessageComposeCreated;

Office.actions.associate("onNewMessageComposeCreated", onNewMessageComposeCreated);
