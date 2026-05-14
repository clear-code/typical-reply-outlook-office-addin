/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2026 ClearCode Inc.
*/
import { ConfigLoader } from "./config-loader.mjs";
import { ReplayMailDataCreator } from "./mail-data-creator.mjs";
import { OfficeDataAccessHelper } from "./office-data-access-helper.mjs";
import { ButtonConfigEnums } from "./config.mjs";

Office.onReady(() => {
  onTypicalReplyButtonClicked();
});

async function singleMailHandler(buttonConfig) {
  console.debug("conversationId: " + Office.context.mailbox.item.conversationId);
  const originalMailData = {
    toRecipients: Office.context.mailbox.item.to.map((recipients) => recipients.emailAddress),
    ccRecipients: Office.context.mailbox.item.cc.map((recipients) => recipients.emailAddress),
    bccRecipients: Office.context.mailbox.item.bcc.map((recipients) => recipients.emailAddress),
    sender: Office.context.mailbox.item.sender?.emailAddress,
    subject: Office.context.mailbox.item.subject,
    id: Office.context.mailbox.item.itemId,
  };
  if (!ReplayMailDataCreator.isAllRecipientsAllowed({ buttonConfig, originalMailData })) {
    console.log("Recipients contains some prohibited domains");
    Office.context.ui.closeContainer();
    return;
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
      Office.context.ui.closeContainer();
    }
  );
}

function groupKeyFor(item) {
  if (!item.conversationId) return null;
  return `${item.conversationId}|${item.subject ?? ""}`;
}

function getDedupeKey(item, canonicalInternetMessageIdByGroup) {
  if (item.internetMessageId) {
    return `imid:${item.internetMessageId}`;
  }
  const gk = groupKeyFor(item);
  if (gk && canonicalInternetMessageIdByGroup?.has(gk)) {
    return `imid:${canonicalInternetMessageIdByGroup.get(gk)}`;
  }
  if (gk) {
    return `conv:${gk}`;
  }
  return `id:${item.itemId}`;
}

async function loadSelectedMails() {
  // As Office addin specification, selected items length is at most 100, so it is safe to load all selected items into memory.
  let selectedItems = await OfficeDataAccessHelper.getSelectedItemsAsync();
  if (selectedItems.length === 0) {
    console.log("No selected items found.");
    return [];
  }
  console.debug(`Selected items count: ${selectedItems.length}`);
  // loadItemByIdAsync must run serially (unloadAsync between loads), so fill
  // in missing internetMessageId / dateTimeCreated one item at a time.
  if (Office.context.requirements.isSetSupported("Mailbox", "1.15")) {
    for (const item of selectedItems) {
      if (!item.itemId) continue;
      if (item.internetMessageId && item.dateTimeCreated) continue;
      const ewsId = Office.context.mailbox.convertToEwsId(
        item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
      );
      const loaded = await OfficeDataAccessHelper.loadItemPropertiesByIdAsync(ewsId);
      if (loaded?.internetMessageId && !item.internetMessageId) {
        item.internetMessageId = loaded.internetMessageId;
      }
    }
  }
  // Collect a representative internetMessageId per (conversationId, subject) group, so siblings
  // "without Message-ID" use it. Note that siblings with Message-ID use its own Message-ID as the key,
  // so they won't be grouped together, that's an intended behavior because if Message-IDs are present,
  // they should be used for grouping.
  const canonicalInternetMessageIdByGroup = new Map();
  for (const item of selectedItems) {
    if (!item.internetMessageId) continue;
    const gk = groupKeyFor(item);
    if (!gk) continue;
    if (!canonicalInternetMessageIdByGroup.has(gk)) {
      canonicalInternetMessageIdByGroup.set(gk, item.internetMessageId);
    }
  }
  const seenDedupeKeys = new Set();
  selectedItems = selectedItems.filter((item) => {
    const key = getDedupeKey(item, canonicalInternetMessageIdByGroup);
    if (seenDedupeKeys.has(key)) {
      return false;
    }
    seenDedupeKeys.add(key);
    return true;
  });
  return selectedItems.map((item) => ({
    toRecipients: item.to,
    ccRecipients: item.cc,
    bccRecipients: item.bcc,
    sender: item.sender?.emailAddress,
    subject: item.subject,
    id: item.itemId,
  }));
}

async function multiMailHandler(buttonConfig) {
  // For multi-select with reading pane, we can not use "reply" or "replay all", we can only create a new mail,
  // and original recipients should not be specified to the new mail recipients because it is insecure.
  if (
    buttonConfig.recipientsType !== ButtonConfigEnums.RecipientsType.SpecifiedByUser &&
    buttonConfig.recipientsType !== ButtonConfigEnums.RecipientsType.Blank
  ) {
    console.log(
      "For multi-select with reading pane, only SpecifiedByUser or Blank recipients type are allowed."
    );
    Office.context.ui.closeContainer();
    return;
  }

  if (!ReplayMailDataCreator.isAllRecipientsAllowed({ buttonConfig, originalMailData: {} })) {
    console.log("Recipients contains some prohibited domains");
    Office.context.ui.closeContainer();
    return;
  }

  const attachments = [];
  if (buttonConfig.forwardType === ButtonConfigEnums.ForwardType.Attachment) {
    const originalMailDataList = await loadSelectedMails();
    if (originalMailDataList.length === 0) {
      console.log("No valid selected mails found.");
      Office.context.ui.closeContainer();
      return;
    }
    for (const originalMailData of originalMailDataList) {
      const attachmentsOfMail = ReplayMailDataCreator.getAttachments({
        buttonConfig,
        originalMailData,
      });
      attachments.push(...attachmentsOfMail);
    }
  }
  const subject = ReplayMailDataCreator.createSubject({ buttonConfig, originalSubject: "" });
  const recipients = ReplayMailDataCreator.getNewRecipients(buttonConfig);
  await OfficeDataAccessHelper.displayNewMessageAsync({
    toRecipients: recipients.to,
    ccRecipients: recipients.cc,
    bccRecipients: recipients.bcc,
    subject: subject,
    htmlBody: buttonConfig.body ? plainTextToHtml(buttonConfig.body) : "",
    attachments,
  });
  Office.context.ui.closeContainer();
}

async function onTypicalReplyButtonClicked() {
  try {
    console.log("onTypicalReplyButtonClicked triggered");
    const params = new URLSearchParams(window.location.search);
    const actionId = params.get("actionId");
    console.debug("actionId:", actionId);
    const buttonConfig = await ConfigLoader.loadButtonConfig(
      Office.context.displayLanguage,
      actionId
    );
    if (!buttonConfig) {
      console.log("No button config find.");
      Office.context.ui.closeContainer();
      return;
    }
    const element = document.getElementById("processing");
    if (buttonConfig.taskPaneMessage) {
      element.innerText = buttonConfig.taskPaneMessage;
    }
    element.hidden = false;
    if (Office.context.mailbox.diagnostics?.hostName === "Outlook") {
      // On Classic Outlook, Office.context.mailbox.item is always defined,
      // regardless of whether one or multiple messages are selected.
      // So we need to check the number of selected items to determine whether
      // it is single-select or multi-select.
      const selectedItems = await loadSelectedMails();
      if (selectedItems.length > 1) {
        await multiMailHandler(buttonConfig);
      } else {
        await singleMailHandler(buttonConfig);
      }
    } else {
      // On Outlook on the web and New Outlook, Office.context.mailbox.item is
      // defined only when a single message is selected.
      // singleMailHandler also requires Office.context.mailbox.item to fetch
      // the original mail data and to display the reply form. Therefore, when
      // Office.context.mailbox.item is undefined, the case must be handled by
      // multiMailHandler, which does not rely on Office.context.mailbox.item.
      const item = Office.context.mailbox.item;
      if (item) {
        await singleMailHandler(buttonConfig);
      } else {
        await multiMailHandler(buttonConfig);
      }
    }
  } catch (e) {
    console.error("onTypicalReplyButtonClicked Failed:", e);
    Office.context.ui.closeContainer();
  }
}

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
