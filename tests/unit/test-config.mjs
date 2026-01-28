/*
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
*/
"use strict";

import { TypicalReplyConfig, ButtonConfig, Config } from "../../src/web/config.mjs";
import { assert } from "tiny-esm-test-runner";
const { is } = assert;

test_TypicalReplyConfig_constructor.parameters = {
  "empty": {
    argument: {
    },
    expected: {
      priority: 0,
      configList: []
    },
  },
  "full": {
    argument: {
      "Priority": 1,
      "ConfigList": [
        {
          "Culture": "ja-JP",
          "GroupLabel": "定型返信",
          "ButtonConfigList": [
            {
              "Id": "button1",
              "Label": "いいね！",
              "Subject": "件名",
              "SubjectPrefix": "[[いいね！]]",
              "Body": "いいね！",
              "Recipients": [
                "test@example.com"
              ],
              "QuoteType": true,
              "AllowedDomains": [
                "*"
              ],
              "ForwardType": "Attachment",
            }
          ]
        },
      ]
    },
    expected: {
      priority: 1,
      configList: [
        {
          culture: "ja-JP",
          groupLabel: "定型返信",
          buttonConfigList: [
            {
              id: "button1",
              label: "いいね！",
              subjectPrefix: "[[いいね！]]",
              subject: "件名",
              body: "いいね！",
              recipients: [
                "test@example.com"
              ],
              recipientsType: 4,
              quoteType: true,
              allowedDomains: ["*"],
              allowedDomainsType: 1,
              forwardType: 1,
            },
          ],
        },
      ],
    }
  }
}
export function test_TypicalReplyConfig_constructor({ argument, expected }) {
  const typicalReplyConfig = new TypicalReplyConfig(argument);
  is(
    typicalReplyConfig,
    expected
  );
}

test_ButtonConfig_constructor.parameters = {
  "empty": {
    argument: {
    },
    expected: {
      id: "",
      label: "",
      subjectPrefix: "",
      subject: "",
      body: "",
      recipients: [],
      recipientsType: 1,
      quoteType: false,
      allowedDomains: [],
      allowedDomainsType: 1,
      forwardType: 0,
    },
  },
  "full": {
    argument: {
      "Id": "button1",
      "Label": "いいね！",
      "Subject": "件名",
      "SubjectPrefix": "[[いいね！]]",
      "Body": "いいね！",
      "Recipients": [
        "test@example.com"
      ],
      "QuoteType": true,
      "AllowedDomains": [
        "*"
      ],
      "ForwardType": "Attachment",
    },
    expected: {
      id: "button1",
      label: "いいね！",
      subjectPrefix: "[[いいね！]]",
      subject: "件名",
      body: "いいね！",
      recipients: [
        "test@example.com"
      ],
      recipientsType: 4,
      quoteType: true,
      allowedDomains: ["*"],
      allowedDomainsType: 1,
      forwardType: 1,
    }
  },
  "parse ForwardType Inline": {
    argument: {
      "ForwardType": "Inline",
    },
    expected: {
      id: "",
      label: "",
      subjectPrefix: "",
      subject: "",
      body: "",
      recipients: [],
      recipientsType: 1,
      quoteType: false,
      allowedDomains: [],
      allowedDomainsType: 1,
      forwardType: 2,
    }
  },
  "parse RecipientsType Blank": {
    argument: {
      "Recipients": [],
    },
    expected: {
      id: "",
      label: "",
      subjectPrefix: "",
      subject: "",
      body: "",
      recipients: [],
      recipientsType: 1,
      quoteType: false,
      allowedDomains: [],
      allowedDomainsType: 1,
      forwardType: 0,
    }
  },
  "parse RecipientsType Sender": {
    argument: {
      "Recipients": ["Sender"],
    },
    expected: {
      id: "",
      label: "",
      subjectPrefix: "",
      subject: "",
      body: "",
      recipients: ["Sender"],
      recipientsType: 2,
      quoteType: false,
      allowedDomains: [],
      allowedDomainsType: 1,
      forwardType: 0,
    }
  },
  "parse RecipientsType All": {
    argument: {
      "Recipients": ["All"],
    },
    expected: {
      id: "",
      label: "",
      subjectPrefix: "",
      subject: "",
      body: "",
      recipients: ["All"],
      recipientsType: 3,
      quoteType: false,
      allowedDomains: [],
      allowedDomainsType: 1,
      forwardType: 0,
    }
  },
  "parse RecipientsType SpecifiedByUser": {
    argument: {
      "Recipients": ["test@example.com"],
    },
    expected: {
      id: "",
      label: "",
      subjectPrefix: "",
      subject: "",
      body: "",
      recipients: ["test@example.com"],
      recipientsType: 4,
      quoteType: false,
      allowedDomains: [],
      allowedDomainsType: 1,
      forwardType: 0,
    }
  },
  "parse AllowedDomainsType All": {
    argument: {
      "AllowedDomains": ["*"],
    },
    expected: {
      id: "",
      label: "",
      subjectPrefix: "",
      subject: "",
      body: "",
      recipients: [],
      recipientsType: 1,
      quoteType: false,
      allowedDomains: ["*"],
      allowedDomainsType: 1,
      forwardType: 0,
    }
  },
  "parse AllowedDomainsType SpecifiedByUser": {
    argument: {
      "AllowedDomains": ["example.com"],
    },
    expected: {
      id: "",
      label: "",
      subjectPrefix: "",
      subject: "",
      body: "",
      recipients: [],
      recipientsType: 1,
      quoteType: false,
      allowedDomains: ["example.com"],
      allowedDomainsType: 2,
      forwardType: 0,
    }
  },
}
export function test_ButtonConfig_constructor({ argument, expected }) {
  const buttonConfig = new ButtonConfig(argument);
  is(
    buttonConfig,
    expected
  );
}