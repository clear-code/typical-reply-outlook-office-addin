/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2025 ClearCode Inc.
*/
import { ButtonConfigEnums } from "./config.mjs";
import * as RecipientParser from "./recipient-parser.mjs";

export class MailDataCreator {
  static CreateDataOnForReplyForm({ config, originalMailData }) {
    const mailItemToReply = {};
    switch (config.RecipientsType) {
      case ButtonConfigEnums.RecipientsType.All:
        mailItemToReply.executeMethod = Office.context.mailbox.item.displayReplyAllFormAsync;
        mailItemToReply.toRecipients = originalMailData.toRecipients;
        mailItemToReply.ccRecipients = originalMailData.ccRecipients;
        mailItemToReply.bccRecipients = originalMailData.bccRecipients;
        break;
      case ButtonConfigEnums.RecipientsType.Sender:
        mailItemToReply.executeMethod = Office.context.mailbox.item.displayReplyFormAsync;
        mailItemToReply.toRecipients = originalMailData.sender;
        break;
      case ButtonConfigEnums.RecipientsType.SpecifiedByUser:
        mailItemToReply.executeMethod = Office.context.mailbox.item.displayReplyFormAsync;
        mailItemToReply.toRecipients = config.Recipients;
        break;
      default:
        break; 
    }
    if (config.AllowedDomainsType == ButtonConfigEnums.AllowedDomainsType.SpecifiedByUser) {
      const loweredAllowedDomains = config.AllowedDomains.toLowerCase();
      for (const recipient of [
        ...(mailItemToReply.toRecipients ?? []),
        ...(mailItemToReply.ccRecipients ?? []),
        ...(mailItemToReply.toRecipients ?? []),
      ]) {
        const parsedRecipient = RecipientParser.parse(recipient);
        if (loweredAllowedDomains.Any((_) => _ == parsedRecipient.domain)) {
          continue;
        }
        console.log(`Prohibited domain: ${parsedRecipient.domain}`);
        return null;
      }
    }
    switch (config.ForwardType) {
      case ButtonConfigEnums.ForwardType.Attachment:
        mailItemToReply.attachments = [
          {
            name: originalMailData.subject,
            type: Office.MailboxEnums.AttachmentType.Item,
            itemId: originalMailData.id,
          },
        ];
        break;
      case ButtonConfigEnums.ForwardType.Inline:
        mailItemToReply.attachments = [
          {
            name: originalMailData.subject,
            type: Office.MailboxEnums.AttachmentType.Item,
            itemId: originalMailData.id,
          },
        ];
        break;
    }
    return mailItemToReply;
  }
  static CreateReplyMailData({ config, originalSuject }) {
    const mailItemToReply = {};
    switch (config.RecipientsType) {
      case ButtonConfigEnums.RecipientsType.SpecifiedByUser:
        mailItemToReply.newToRecipients = config.Recipients;
        break;
      default:
        break;
    }

    if (config.Subject) {
      mailItemToReply.subject = config.Subject;
    }
    else {
      mailItemToReply.subject = originalSuject;
    }

    if (config.SubjectPrefix) {
      mailItemToReply.subject = `${config.SubjectPrefix} ${mailItemToReply.subject ?? ""}`;
    }

    mailItemToReply.bodyHtml = config.Body ?? "";
    mailItemToReply.quoteType = config.QuoteType
    return mailItemToReply;
  }
}
