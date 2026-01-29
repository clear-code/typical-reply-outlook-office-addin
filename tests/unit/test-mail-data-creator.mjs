/*
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/
"use strict";

import { ButtonConfigEnums } from "../../src/web/config.mjs";
import { ReplayMailDataCreator } from "../../src/web/mail-data-creator.mjs";
import { assert } from "tiny-esm-test-runner";
import { OfficeMockObject } from 'office-addin-mock';
const { is } = assert;

const mockData = {
  host: "outlook", // Outlookの場合必須
  context: {
    mailbox: {
      item: {
        displayReplyAllFormAsync: function () { },
        displayReplyFormAsync: function () { }
      }
    }
  },
  MailboxEnums: {
    AttachmentType: {
      Cloud: "cloud",
      File: "file",
      Item: "item",
      Base64: "base64",
    }
  }
};
const officeMock = new OfficeMockObject(mockData);
global.Office = officeMock;

test_getReplyFormFunction.parameters = {
  "empty": {
    buttonConfig: {
    },
    expectedFuncName: "displayReplyFormAsync"
  },
  "RecipientsType.All": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.All
    },
    expectedFuncName: "displayReplyAllFormAsync"
  },
  "RecipientsType.Sender": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.Sender
    },
    expectedFuncName: "displayReplyFormAsync"
  },
  "RecipientsType.SpecifiedByUser": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.SpecifiedByUser
    },
    expectedFuncName: "displayReplyFormAsync"
  },
}
export function test_getReplyFormFunction({ buttonConfig, expectedFuncName }) {
  const func = ReplayMailDataCreator.getReplyFormFunction(buttonConfig);
  is(
    expectedFuncName,
    func.name
  );
}

test_isAllRecipientsAllowed.parameters = {
  "emptySetting": {
    buttonConfig: {
    },
    originalMailData: {
    },
    expected: true,
  },
  "All": {
    buttonConfig: {
      allowedDomainsType: ButtonConfigEnums.AllowedDomainsType.All,
    },
    originalMailData: {
      toRecipients: ["test@to.example.com"],
      ccRecipients: ["test@cc.example.com"],
      bccRecipients: ["test@bcc.example.com"]
    },
    expected: true,
  },
  "RecipientsType.All and AllowedDomainsType.SpecifiedByUser accepted": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.All,
      allowedDomains: [
        "to.example.com",
        "cc.example.com",
        "bcc.example.com",
        "sender.example.com",
      ],
      allowedDomainsType: ButtonConfigEnums.AllowedDomainsType.SpecifiedByUser,
    },
    originalMailData: {
      toRecipients: ["test@to.example.com"],
      ccRecipients: ["test@cc.example.com"],
      bccRecipients: ["test@bcc.example.com"],
      sender: "test@sender.example.com"
    },
    expected: true,
  },
  "RecipientsType.All and AllowedDomainsType.SpecifiedByUser rejected": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.All,
      allowedDomains: [
        "to.example.com",
        "cc.example.com",
        "bcc.example.com",
      ],
      allowedDomainsType: ButtonConfigEnums.AllowedDomainsType.SpecifiedByUser,
    },
    originalMailData: {
      toRecipients: ["test@to.example.com"],
      ccRecipients: ["test@cc.example.com"],
      bccRecipients: ["test@bcc.example.com"],
      sender: "test@sender.example.com"
    },
    expected: false,
  },
  "RecipientsType.Sender and AllowedDomainsType.SpecifiedByUser accepted": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.Sender,
      allowedDomains: [
        "sender.example.com",
      ],
      allowedDomainsType: ButtonConfigEnums.AllowedDomainsType.SpecifiedByUser,
    },
    originalMailData: {
      sender: "test@sender.example.com"
    },
    expected: true,
  },
  "RecipientsType.Sender and AllowedDomainsType.SpecifiedByUser rejected": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.Sender,
      allowedDomains: [
        "to.example.com",
      ],
      allowedDomainsType: ButtonConfigEnums.AllowedDomainsType.SpecifiedByUser,
    },
    originalMailData: {
      sender: "test@sender.example.com"
    },
    expected: false,
  },
  "RecipientsType.SpecifiedByUser and AllowedDomainsType.SpecifiedByUser accepted": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.SpecifiedByUser,
      recipients: [
        "test@to.example.com",
        "test@to2.example.com",
        "test@to3.example.com"
      ],
      allowedDomains: [
        "to.example.com",
        "to2.example.com",
        "to3.example.com",
      ],
      allowedDomainsType: ButtonConfigEnums.AllowedDomainsType.SpecifiedByUser,
    },
    originalMailData: {
    },
    expected: true,
  },
  "RecipientsType.SpecifiedByUser and AllowedDomainsType.SpecifiedByUser rejected": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.SpecifiedByUser,
      recipients: [
        "test@to.example.com",
        "test@to2.example.com",
        "test@to3.example.com"
      ],
      allowedDomains: [
        "to.example.com",
        "to2.example.com",
      ],
      allowedDomainsType: ButtonConfigEnums.AllowedDomainsType.SpecifiedByUser,
    },
    originalMailData: {
    },
    expected: false,
  },
}
export function test_isAllRecipientsAllowed({ buttonConfig, originalMailData, expected }) {
  const result = ReplayMailDataCreator.isAllRecipientsAllowed({ buttonConfig, originalMailData });
  is(
    expected,
    result
  );
}

test_getAttachments.parameters = {
  "emptySetting": {
    buttonConfig: {
    },
    originalMailData: {
    },
    expected: [],
  },
  "ForwardType.Attachment": {
    buttonConfig: {
      forwardType: ButtonConfigEnums.ForwardType.Attachment,
    },
    originalMailData: {
      id: "originalMailId",
      attachments: [
        { id: "att1" }
      ],
      subject: "Original Subject",
    },
    expected: [
      {
        name: "Original Subject",
        type: Office.MailboxEnums.AttachmentType.Item,
        itemId: "originalMailId",
      }
    ],
  },
  "Empty subject": {
    buttonConfig: {
      forwardType: ButtonConfigEnums.ForwardType.Attachment
    },
    originalMailData: {
      id: "originalMailId",
      attachments: [
        { id: "att1" }
      ],
    },
    expected: [
      {
        name: " ",
        type: Office.MailboxEnums.AttachmentType.Item,
        itemId: "originalMailId",
      }
    ],
  },
  "ForwardType.Inline": {
    buttonConfig: {
      forwardType: ButtonConfigEnums.ForwardType.Inline
    },
    originalMailData: {
      id: "originalMailId",
      attachments: [
        { id: "att1" }
      ],
      subject: "Original Subject",
    },
    expected: [
      {
        name: "Original Subject",
        type: Office.MailboxEnums.AttachmentType.Item,
        itemId: "originalMailId",
      }
    ],
  },
  "ForwardType.Unknown": {
    buttonConfig: {
      forwardType: ButtonConfigEnums.ForwardType.Unknown
    },
    originalMailData: {
      id: "originalMailId",
      attachments: [
        { id: "att1" }
      ],
      subject: "Original Subject",
    },
    expected: [],
  },
}
export function test_getAttachments({ buttonConfig, originalMailData, expected }) {
  const result = ReplayMailDataCreator.getAttachments({ buttonConfig, originalMailData });
  is(
    expected,
    result
  );
}

test_getNewRecipients.parameters = {
  "emptySetting": {
    buttonConfig: {
    },
    expected: {},
  },
  "RecipientsType.SpecifiedByUser": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.SpecifiedByUser,
      recipients: [
        "test@example.com",
        "test2@example.com"
      ],
    },
    expected: {
      to: [
        "test@example.com",
        "test2@example.com"
      ],
      cc: [],
      bcc: [],
    },
  },
  "RecipientsType.All": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.All,
      recipients: [
        "test@example.com",
        "test2@example.com"
      ],
    },
    expected: {
    },
  },
  "RecipientsType.Sender": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.Sender,
      recipients: [
        "test@example.com",
        "test2@example.com"
      ],
    },
    expected: {
    },
  },
  "RecipientsType.Blank": {
    buttonConfig: {
      recipientsType: ButtonConfigEnums.RecipientsType.Blank,
      recipients: [
        "test@example.com",
        "test2@example.com"
      ],
    },
    expected: {
      to: [],
      cc: [],
      bcc: [],
    },
  },
}
export function test_getNewRecipients({ buttonConfig, expected }) {
  const result = ReplayMailDataCreator.getNewRecipients(buttonConfig);
  is(
    expected,
    result
  );
}

test_createSubject.parameters = {
  "empty": {
    buttonConfig: {
    },
    originalSubject: "",
    expected: "",
  },
  "empty config": {
    buttonConfig: {
    },
    originalSubject: "Original Subject",
    expected: "Original Subject",
  },
  "subject prefix": {
    buttonConfig: {
      subjectPrefix: "[Prefix]",
    },
    originalSubject: "Original Subject",
    expected: "[Prefix] Original Subject",
  },
  "only subject prefix": {
    buttonConfig: {
      subjectPrefix: "[Prefix]",
    },
    originalSubject: "Original Subject",
    expected: "[Prefix] Original Subject",
  },
  "subject and subject prefix": {
    buttonConfig: {
      subjectPrefix: "[Prefix]",
      subject: "New Subject",
    },
    originalSubject: "Original Subject",
    expected: "[Prefix] New Subject",
  },
}
export function test_createSubject({ buttonConfig, originalSubject, expected }) {
  const result = ReplayMailDataCreator.createSubject({ buttonConfig, originalSubject });
  is(
    expected,
    result
  );
}
