/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2025 ClearCode Inc.
*/
export class Config {
  Culture;
  GroupLabel = "Typical Reply";
  TabMailInsertAfterMso = "GroupMailRespond";
  TabReadInsertAfterMso = "GroupRespond";
  ContextMenuInsertAfterMso = "Forward";
  ButtonConfigList;

  constructor({
    Culture,
    GroupLabel,
    TabMailInsertAfterMso,
    TabReadInsertAfterMso,
    ContextMenuInsertAfterMso,
    ButtonConfigList,
  }) {
    this.Culture = Culture;
    this.GroupLabel = GroupLabel;
    this.TabMailInsertAfterMso = TabMailInsertAfterMso;
    this.TabReadInsertAfterMso = TabReadInsertAfterMso;
    this.ContextMenuInsertAfterMso = ContextMenuInsertAfterMso;
    this.ButtonConfigList = ButtonConfigList;
  }
}

const ForwardType = {
  Unknown: 0,
  Attachment: 1,
  Inline: 2,
};

const RecipientsType = {
  Unknown: 0,
  Blank: 1,
  Sender: 2,
  All: 3,
  SpecifiedByUser: 4,
};

const AllowedDomainsType = {
  Unknown: 0,
  All: 1,
  SpecifiedByUser: 2,
};

const ButtonSize = {
  Unknown: 0,
  Normal: 1,
  Large: 2,
};

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
  Size;
  Image = "logo.png";

  constructor({
    Id,
    Label,
    SubjectPrefix,
    Subject,
    Body,
    Recipients,
    QuoteType,
    AllowedDomains,
    AllowedDomainsType,
    ForwardType,
    Size,
    Image,
  }) {
    this.Id = Id;
    this.Label = Label;
    this.SubjectPrefix = SubjectPrefix;
    this.Subject = Subject;
    this.Body = Body;
    this.Recipients = Recipients;
    this.QuoteType = QuoteType;
    this.AllowedDomains = AllowedDomains;
    this.ForwardType = ForwardType;
    this.Size = Size;
    this.Image = Image;

    if (!Recipients || Recipients.length == 0) {
      this.RecipientsType = RecipientsType.Blank;
    } else {
      this.RecipientsType = RecipientsType.SpecifiedByUser;
      for (const key of Object.keys(RecipientsType)) {
        const lowerdKey = key.toLowerCase();
        const inputRecipientLower = Recipients[0].toLowerCase();
        if (lowerdKey === inputRecipientLower) {
          this.RecipientsType = RecipientsType[key];
          break;
        }
      }
    }

    if (!AllowedDomains || AllowedDomains.length == 0 || AllowedDomains[0] === "*") {
      this.AllowedDomainsType = AllowedDomainsType.Blank;
    } else {
      this.AllowedDomainsType = AllowedDomainsType.SpecifiedByUser;
    }
  }
}
