/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2025 ClearCode Inc.
*/
import * as RecipientParser from "./recipient-parser.mjs";

export class OfficeDataAccessHelper {
  static getBccAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.bcc.getAsync((asyncResult) => {
          const recipients = asyncResult.value.map((officeAddonRecipient) => ({
            ...officeAddonRecipient,
            ...RecipientParser.parse(officeAddonRecipient.emailAddress),
          }));
          resolve(recipients);
        });
      } catch (error) {
        console.log(`Error while getting Bcc: ${error}`);
        reject(error);
      }
    });
  }

  static setBccAsync(recipients) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.bcc.setAsync(recipients, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(recipients);
          } else {
            console.log(`Error while setting Bcc: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
        });
      } catch (error) {
        console.log(`Error while setting Bcc: ${error}`);
        reject(error);
      }
    });
  }

  static getCcAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.cc.getAsync((asyncResult) => {
          const recipients = asyncResult.value.map((officeAddonRecipient) => ({
            ...officeAddonRecipient,
            ...RecipientParser.parse(officeAddonRecipient.emailAddress),
          }));
          resolve(recipients);
        });
      } catch (error) {
        console.log(`Error while getting Cc: ${error}`);
        reject(error);
      }
    });
  }

  static setCcAsync(recipients) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.cc.setAsync(recipients, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(recipients);
          } else {
            console.log(`Error while setting Cc: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
        });
      } catch (error) {
        console.log(`Error while setting Cc: ${error}`);
        reject(error);
      }
    });
  }

  static clearCcAsync() {
    return OfficeDataAccessHelper.setCcAsync([]);
  }

  static getSubjectAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.subject.getAsync((asyncResult) => {
          const subject = asyncResult.value;
          resolve(subject);
        });
      } catch (error) {
        console.log(`Error while getting subject: ${error}`);
        reject(error);
      }
    });
  }

  static getBodyAsync(coerctionType = Office.CoercionType.Html) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.body.getAsync(
          coerctionType,
          { bodyMode: Office.MailboxEnums.BodyMode.Full },
          (asyncResult) => {
            const body = asyncResult.value;
            resolve(body);
          }
        );
      } catch (error) {
        console.log(`Error while getting body: ${error}`);
        reject(error);
      }
    });
  }

  static getItemIdAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.getItemIdAsync((asyncResult) => {
          const id = asyncResult.value;
          resolve(id);
        });
      } catch (error) {
        console.log(`Error while getting itemId: ${error}`);
        reject(error);
      }
    });
  }

  static loadCustomPropertiesAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.loadCustomPropertiesAsync((asyncResult) => {
          resolve(asyncResult.value);
        });
      } catch (error) {
        console.log(`Error while getting itemId: ${error}`);
        reject(error);
      }
    });
  }

  static getRequiredAttendeeAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.requiredAttendees.getAsync((asyncResult) => {
          const recipients = asyncResult.value.map((officeAddonRecipient) => ({
            ...officeAddonRecipient,
            ...RecipientParser.parse(officeAddonRecipient.emailAddress),
          }));
          resolve(recipients);
        });
      } catch (error) {
        console.log(`Error while getting required attendees: ${error}`);
        reject(error);
      }
    });
  }

  static getOptionalAttendeeAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.optionalAttendees.getAsync((asyncResult) => {
          const recipients = asyncResult.value.map((officeAddonRecipient) => ({
            ...officeAddonRecipient,
            ...RecipientParser.parse(officeAddonRecipient.emailAddress),
          }));
          resolve(recipients);
        });
      } catch (error) {
        console.log(`Error while getting optional attendees: ${error}`);
        reject(error);
      }
    });
  }

  static getToAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.to.getAsync((asyncResult) => {
          const recipients = asyncResult.value.map((officeAddonRecipient) => ({
            ...officeAddonRecipient,
            ...RecipientParser.parse(officeAddonRecipient.emailAddress),
          }));
          resolve(recipients);
        });
      } catch (error) {
        console.log(`Error while getting To: ${error}`);
        reject(error);
      }
    });
  }

  static setToAsync(recipients) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.to.setAsync(recipients, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(recipients);
          } else {
            console.log(`Error while setting To: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
        });
      } catch (error) {
        console.log(`Error while setting To: ${error}`);
        reject(error);
      }
    });
  }

  static clearToAsync() {
    return OfficeDataAccessHelper.setToAsync([]);
  }

  static getSessionDataAsync(key) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.sessionData.getAsync(key, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(asyncResult.value);
          } else {
            console.debug(`Error while getting SessionData [${key}]: ${asyncResult.error.message}`);
            // Regards no value
            resolve("");
          }
        });
      } catch (error) {
        console.log(`Error while getting SessionData [${key}]: ${error}`);
        reject(error);
      }
    });
  }

  static getDelayDeliveryTime() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.delayDeliveryTime.getAsync((asyncResult) => {
          const value = asyncResult.value;
          resolve(value);
        });
      } catch (error) {
        console.log(`Error while getting DelayDeliveryTime: ${error}`);
        reject(error);
      }
    });
  }

  static getInitializationContextAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.getInitializationContextAsync((asyncResult) => {
          const value = asyncResult.value.itemId;
          resolve(value);
        });
      } catch (error) {
        console.log(`Error while getting getInitializationContextAsync: ${error}`);
        reject(error);
      }
    });
  }

  static setDelayDeliveryTimeAsync(deliveryTime) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.delayDeliveryTime.setAsync(deliveryTime, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            resolve(false);
          } else {
            resolve(true);
          }
        });
      } catch (error) {
        console.log(`Error while setting DelayDeliveryTime: ${error}`);
        reject(error);
      }
    });
  }

  static setSubjectAsync(subject) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.subject.setAsync(subject, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(true);
          } else {
            console.log(`Error while setting subject: ${asyncResult.error.message}`);
            reject(false);
          }
        });
      } catch (error) {
        console.log(`Error while setting subject: ${error}`);
        reject(error);
      }
    });
  }

  static setBodyAsync(body) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.body.setAsync(body, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(true);
          } else {
            console.log(`Error while setting body: ${asyncResult.error.message}`);
            reject(false);
          }
        });
      } catch (error) {
        console.log(`Error while setting body: ${error}`);
        reject(error);
      }
    });
  }

  static prependBodyAsync(body) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.body.prependAsync(body, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(true);
          } else {
            console.log(`Error while prepending body: ${asyncResult.error.message}`);
            reject(false);
          }
        });
      } catch (error) {
        console.log(`Error while prepending body: ${error}`);
        reject(error);
      }
    });
  }

  static saveRoamingSettingsAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.roamingSettings.saveAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(true);
          } else {
            console.log(`Error while saving RoamingSettings: ${asyncResult.error.message}`);
            reject(false);
          }
        });
      } catch (error) {
        console.log(`Error while saving RoamingSettings: ${error}`);
        reject(error);
      }
    });
  }
}
