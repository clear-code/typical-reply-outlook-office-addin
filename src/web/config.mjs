/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2025 ClearCode Inc.
*/

function getEnumValueByKey(enumObj, key) {
  if (!key) {
    return undefined;
  }
  const lowerdKey = key.toLowerCase();
  for (const enumKey in enumObj) {
    if (enumKey.toLowerCase() === lowerdKey) {
      return enumObj[enumKey];
    }
  }
  return undefined;
}

export class ButtonConfigEnums {
  static ForwardType = {
    Unknown: 0,
    Attachment: 1,
    Inline: 2,
  };

  static RecipientsType = {
    Unknown: 0,
    Blank: 1,
    Sender: 2,
    All: 3,
    SpecifiedByUser: 4,
  };

  static AllowedDomainsType = {
    Unknown: 0,
    All: 1,
    SpecifiedByUser: 2,
  };
}

export class ButtonConfig {
  id;
  label;
  subjectPrefix;
  subject;
  body;
  recipients;
  recipientsType;
  quoteType;
  allowedDomains;
  allowedDomainsType;
  forwardType;

  constructor({
    Id,
    Label,
    SubjectPrefix,
    Subject,
    Body,
    Recipients,
    QuoteType,
    AllowedDomains,
    ForwardType,
  }) {
    this.id = Id ?? "";
    this.label = Label ?? "";
    this.subjectPrefix = SubjectPrefix ?? "";
    this.subject = Subject ?? "";
    this.body = Body ?? "";
    this.recipients = Recipients ?? [];
    this.quoteType = QuoteType ?? false;
    this.allowedDomains = AllowedDomains ?? [];
    this.forwardType =
      getEnumValueByKey(ButtonConfigEnums.ForwardType, ForwardType) ??
      ButtonConfigEnums.ForwardType.Unknown;

    if (!Recipients || Recipients.length == 0) {
      this.recipientsType = ButtonConfigEnums.RecipientsType.Blank;
    } else {
      this.recipientsType =
        getEnumValueByKey(ButtonConfigEnums.RecipientsType, Recipients[0].toLowerCase()) ??
        ButtonConfigEnums.RecipientsType.SpecifiedByUser;
    }

    if (!AllowedDomains || AllowedDomains.length == 0 || AllowedDomains[0] === "*") {
      this.allowedDomainsType = ButtonConfigEnums.AllowedDomainsType.All;
    } else {
      this.allowedDomainsType = ButtonConfigEnums.AllowedDomainsType.SpecifiedByUser;
    }
  }
}

export class Config {
  culture;
  groupLabel;
  buttonConfigList;

  constructor({ Culture, GroupLabel, ButtonConfigList }) {
    this.culture = Culture ?? "en-US";
    this.groupLabel = GroupLabel ?? "Typical Reply";
    this.buttonConfigList = [];
    if (ButtonConfigList) {
      for (const buttonConfig of ButtonConfigList) {
        this.buttonConfigList.push(new ButtonConfig(buttonConfig));
      }
    }
  }
}

export class TypicalReplyConfig {
  priority;
  configList;
  constructor({ Priority, ConfigList }) {
    this.priority = Priority ?? 0;
    this.configList = [];
    if (ConfigList) {
      for (const config of ConfigList) {
        this.configList.push(new Config(config));
      }
    }
  }
}
