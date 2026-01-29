/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2025 ClearCode Inc.
*/
import { ButtonConfigEnums } from "./config.mjs";
import * as RecipientParser from "./recipient-parser.mjs";

export class ReplayMailDataCreator {
  static getReplyFormFunction(buttonConfig) {
    switch (buttonConfig.recipientsType) {
      case ButtonConfigEnums.RecipientsType.All:
        return Office.context.mailbox.item.displayReplyAllFormAsync;
      case ButtonConfigEnums.RecipientsType.Sender:
        return Office.context.mailbox.item.displayReplyFormAsync;
      case ButtonConfigEnums.RecipientsType.SpecifiedByUser:
        return Office.context.mailbox.item.displayReplyFormAsync;
      default:
        return Office.context.mailbox.item.displayReplyFormAsync;
    }
  }

  static isAllRecipientsAllowed({ buttonConfig, originalMailData }) {
    let recipients;
    switch (buttonConfig.recipientsType) {
      case ButtonConfigEnums.RecipientsType.All:
        recipients = [
          ...(originalMailData.toRecipients ?? []),
          ...(originalMailData.ccRecipients ?? []),
          ...(originalMailData.bccRecipients ?? []),
          originalMailData.sender,
        ];
        break;
      case ButtonConfigEnums.RecipientsType.Sender:
        if (originalMailData.sender) {
          recipients = [originalMailData.sender];
        }
        break;
      case ButtonConfigEnums.RecipientsType.SpecifiedByUser:
        recipients = buttonConfig.recipients ?? [];
        break;
      default:
        break;
    }
    if (buttonConfig.allowedDomainsType == ButtonConfigEnums.AllowedDomainsType.SpecifiedByUser) {
      const loweredAllowedDomains = buttonConfig.allowedDomains.map((domain) =>
        domain.toLowerCase()
      );
      for (const recipient of recipients) {
        if (!recipient) {
          continue;
        }
        const parsedRecipient = RecipientParser.parse(recipient);
        if (loweredAllowedDomains.find((domain) => domain === parsedRecipient.domain)) {
          continue;
        }
        console.log(`Prohibited domain: ${parsedRecipient.domain}`);
        return false;
      }
    }
    return true;
  }

  static getAttachments({ buttonConfig, originalMailData }) {
    if (!originalMailData.id) {
      return [];
    }
    switch (buttonConfig.forwardType) {
      case ButtonConfigEnums.ForwardType.Attachment:
        return [
          {
            name: originalMailData.subject ?? " ",
            type: Office.MailboxEnums.AttachmentType.Item,
            itemId: originalMailData.id,
          },
        ];
      case ButtonConfigEnums.ForwardType.Inline:
        // TODO: Suport Inline mode
        return [
          {
            name: originalMailData.subject ?? " ",
            type: Office.MailboxEnums.AttachmentType.Item,
            itemId: originalMailData.id,
          },
        ];
    }
    return [];
  }

  static getNewRecipients(buttonConfig) {
    switch (buttonConfig.recipientsType) {
      case ButtonConfigEnums.RecipientsType.SpecifiedByUser:
        return {
          to: buttonConfig.recipients,
          cc: [],
          bcc: [],
        };
      case ButtonConfigEnums.RecipientsType.Blank:
        return {
          to: [],
          cc: [],
          bcc: [],
        };
      default:
        return {};
    }
  }

  static createSubject({ buttonConfig, originalSubject }) {
    let prefix = buttonConfig.subjectPrefix ?
     `${buttonConfig.subjectPrefix} ` :
     "";
    return buttonConfig.subject ? 
      prefix + buttonConfig.subject:
      prefix + originalSubject;
  }
}
