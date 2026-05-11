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

const _runtimeId = Math.random().toString(36).slice(2, 8);
console.log("app.js module evaluated, runtimeId:", _runtimeId);

const TYPICALREPLY_DEDUPE_KEY = "typicalReply.lastInvocationAt";
const TYPICALREPLY_DEDUPE_WINDOW_MS = 2000;

function tryClaimInvocation() {
  try {
    const last = parseInt(localStorage.getItem(TYPICALREPLY_DEDUPE_KEY) || "0", 10);
    const now = Date.now();
    if (now - last < TYPICALREPLY_DEDUPE_WINDOW_MS) {
      return false;
    }
    localStorage.setItem(TYPICALREPLY_DEDUPE_KEY, String(now));
    return true;
  } catch {
    return true;
  }
}

Office.onReady(() => {
  if (!tryClaimInvocation()) {
    console.log("Skipped duplicate onTypicalReplyButtonClicked invocation, runtimeId:", _runtimeId);
    return;
  }
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

function getDedupeKey(item) {
  if (item.internetMessageId) {
    return `imid:${item.internetMessageId}`;
  }
  if (item.conversationId && item.dateTimeCreated) {
    const created =
      item.dateTimeCreated instanceof Date
        ? item.dateTimeCreated.toISOString()
        : String(item.dateTimeCreated);
    return `conv:${item.conversationId}|created:${created}`;
  }
  return `id:${item.itemId}`;
}

async function loadSelectedMails() {
  let selectedItems = await OfficeDataAccessHelper.getSelectedItemsAsync();
  if (selectedItems == null || selectedItems.length === 0) {
    console.log("No selected items found.");
    return null;
  }
  if (selectedItems.length > 100) {
    console.log("Too many selected items.");
    return null;
  }
  // loadItemByIdAsync must run serially (unloadAsync between loads), so fill
  // in missing internetMessageId / dateTimeCreated one item at a time.
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
    if (loaded?.dateTimeCreated && !item.dateTimeCreated) {
      item.dateTimeCreated = loaded.dateTimeCreated;
    }
  }
  const seenDedupeKeys = new Set();
  selectedItems = selectedItems.filter((item) => {
    const key = getDedupeKey(item);
    console.log(key);
    if (seenDedupeKeys.has(key))
    {
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
  if (buttonConfig.recipientsType !== ButtonConfigEnums.RecipientsType.SpecifiedByUser) {
    console.log(
      "For multi-select with reading pane, only SpecifiedByUser recipients type is allowed."
    );
    Office.context.ui.closeContainer();
    return;
  }

  const originalMailDataList = await loadSelectedMails();
  if (!originalMailDataList || originalMailDataList.length === 0) {
    console.log("No valid selected mails found.");
    Office.context.ui.closeContainer();
    return;
  }
  const attachments = [];
  for (const originalMailData of originalMailDataList) {
    if (!ReplayMailDataCreator.isAllRecipientsAllowed({ buttonConfig, originalMailData })) {
      console.log("Recipients contains some prohibited domains");
      Office.context.ui.closeContainer();
      return;
    }
    const attachmentsOfMail = ReplayMailDataCreator.getAttachments({
      buttonConfig,
      originalMailData,
    });
    attachments.push(...attachmentsOfMail);
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
    const item = Office.context.mailbox.item;
    if (item) {
      // No reading pane item (e.g. multi-select with no preview) — nothing to reply to
      await singleMailHandler(buttonConfig);
    } else {
      await multiMailHandler(buttonConfig);
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
