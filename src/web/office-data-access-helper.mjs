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
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const recipients = asyncResult.value.map((officeAddonRecipient) => ({
              ...officeAddonRecipient,
              ...RecipientParser.parse(officeAddonRecipient.emailAddress),
            }));
            resolve(recipients);
          } else {
            console.log(`Error while getting Bcc: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
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
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const recipients = asyncResult.value.map((officeAddonRecipient) => ({
              ...officeAddonRecipient,
              ...RecipientParser.parse(officeAddonRecipient.emailAddress),
            }));
            resolve(recipients);
          } else {
            console.log(`Error while getting Cc: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
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

  static getSubjectAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.subject.getAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const subject = asyncResult.value;
            resolve(subject);
          } else {
            console.log(`Error while getting subject: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
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
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              const body = asyncResult.value;
              resolve(body);
            } else {
              console.log(`Error while getting body: ${asyncResult.error.message}`);
              reject(asyncResult.error);
            }
          }
        );
      } catch (error) {
        console.log(`Error while getting body: ${error}`);
        reject(error);
      }
    });
  }

  static getToAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.to.getAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const recipients = asyncResult.value.map((officeAddonRecipient) => ({
              ...officeAddonRecipient,
              ...RecipientParser.parse(officeAddonRecipient.emailAddress),
            }));
            resolve(recipients);
          } else {
            console.log(`Error while getting To: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
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

  static setSubjectAsync(subject) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.subject.setAsync(subject, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(true);
          } else {
            console.log(`Error while setting subject: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
        });
      } catch (error) {
        console.log(`Error while setting subject: ${error}`);
        reject(error);
      }
    });
  }

  static getBodyTypeAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.body.getTypeAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(asyncResult.value);
          } else {
            console.log(`Error while getting body type: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
        });
      } catch (error) {
        console.log(`Error while getting body type: ${error}`);
        reject(error);
      }
    });
  }

  static setBodyAsync(body, coercionType = Office.CoercionType.Html) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.body.setAsync(body, { coercionType }, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(true);
          } else {
            console.log(`Error while setting body: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
        });
      } catch (error) {
        console.log(`Error while setting body: ${error}`);
        reject(error);
      }
    });
  }

  static prependBodyAsync(body, coercionType = Office.CoercionType.Html) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.item.body.prependAsync(body, { coercionType }, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(true);
          } else {
            console.log(`Error while prepending body: ${asyncResult.error.message}`);
            reject(asyncResult.error);
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
            reject(asyncResult.error);
          }
        });
      } catch (error) {
        console.log(`Error while saving RoamingSettings: ${error}`);
        reject(error);
      }
    });
  }

  static getSelectedItemsAsync() {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            if (!asyncResult.value || asyncResult.value.length === 0) {
              console.debug("No items are selected");
              resolve([]);
              return;
            }
            resolve(asyncResult.value);
          } else {
            console.log(`Error while getting selected items: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
        });
      } catch (error) {
        console.log(`Error while getting selected items: ${error}`);
        reject(error);
      }
    });
  }

  static displayNewMessageAsync(parameters) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.displayNewMessageFormAsync(parameters, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(true);
          } else {
            console.log(`Error while displaying new message: ${asyncResult.error.message}`);
            reject(asyncResult.error);
          }
        });
      } catch (error) {
        console.log(`Error while displaying new message: ${error}`);
        reject(error);
      }
    });
  }

  static loadItemPropertiesByIdAsync(itemId) {
    return new Promise((resolve) => {
      try {
        Office.context.mailbox.loadItemByIdAsync(itemId, (loadResult) => {
          if (loadResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.log(`Error while loading item ${itemId}: ${loadResult.error?.message}`);
            return resolve(null);
          }
          const loaded = loadResult.value;
          const properties = {
            internetMessageId: loaded.internetMessageId || "",
          };
          loaded.unloadAsync((unloadResult) => {
            if (unloadResult.status !== Office.AsyncResultStatus.Succeeded) {
              console.log(`Error while unloading item ${itemId}: ${unloadResult.error?.message}`);
            }
            resolve(properties);
          });
        });
      } catch (error) {
        console.log(`Error while loading item ${itemId}: ${error}`);
        resolve(null);
      }
    });
  }
}
