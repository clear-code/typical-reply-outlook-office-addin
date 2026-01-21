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
  Id;
  Label;
  SubjectPrefix;
  Subject;
  Body;
  Recipients;
  RecipientsType;
  QuoteType;
  AllowedDomains;
  LoweredAllowedDomains;
  AllowedDomainsType;
  ForwardType;

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
    this.Id = Id ?? "";
    this.Label = Label ?? "";
    this.SubjectPrefix = SubjectPrefix ?? "";
    this.Subject = Subject ?? "";
    this.Body = Body ?? "";
    this.Recipients = Recipients ?? [];
    this.QuoteType = QuoteType ?? false;
    this.AllowedDomains = AllowedDomains ?? [];
    this.ForwardType =
      getEnumValueByKey(ButtonConfigEnums.ForwardType, ForwardType) ??
      ButtonConfigEnums.ForwardType.Unknown;

    if (!Recipients || Recipients.length == 0) {
      this.RecipientsType = ButtonConfigEnums.RecipientsType.Blank;
    } else {
      this.RecipientsType = getEnumValueByKey(ButtonConfigEnums.RecipientsType, Recipients[0].toLowerCase()) 
        ?? ButtonConfigEnums.RecipientsType.SpecifiedByUser;
    }

    if (!AllowedDomains || AllowedDomains.length == 0 || AllowedDomains[0] === "*") {
      this.AllowedDomainsType = ButtonConfigEnums.AllowedDomainsType.Blank;
    } else {
      this.AllowedDomainsType = ButtonConfigEnums.AllowedDomainsType.SpecifiedByUser;
    }
  }
}

export class Config {
  Culture;
  GroupLabel;
  ButtonConfigList;

  constructor({ Culture, GroupLabel, ButtonConfigList }) {
    this.Culture = Culture ?? "en-US";
    this.GroupLabel = GroupLabel ?? "Typical Reply";
    this.ButtonConfigList = [];
    if (ButtonConfigList) {
      for (const buttonConfig of ButtonConfigList) {
        this.ButtonConfigList.push(new ButtonConfig(buttonConfig));
      }
    }
  }
}

export class TypicalReplyConfig {
  Priority;
  ConfigList;
  constructor({ Priority, ConfigList }) {
    this.Priority = Priority ?? 0;
    this.ConfigList = [];
    if (ConfigList) {
      for (const config of ConfigList) {
        this.ConfigList.push(new Config(config));
      }
    }
  }
}
