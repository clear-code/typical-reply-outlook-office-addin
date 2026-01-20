/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2025 ClearCode Inc.
*/

import { TypicalReplyConfig } from "./config.mjs";

export class ConfigLoader {
  static async loadFile(url) {
    console.debug("loadFile ", url);
    try {
      const response = await fetch(url, { cache: "no-store" });
      console.debug("response:", response);
      if (!response.ok) {
        return "";
      }
      const data = await response.text();
      return data.trim();
    } catch (err) {
      console.error(err);
      return "";
    }
  }

  static async loadFileConfig() {
    const configJsonString = await this.loadFile("configs/TypicalReplyConfig.json");
    if (configJsonString) {
      const configObject = JSON.parse(configJsonString);
      return new TypicalReplyConfig(configObject);
    }
    return new TypicalReplyConfig({});
  }

  static async loadConfigForCurrentLanguage(culture) {
    const typicalReplyConfig = await ConfigLoader.loadFileConfig();
    let config = typicalReplyConfig?.ConfigList?.find((_) => (_.Culture ?? null) === culture);
    if (!config) {
      const lang = culture.split("-")[0];
      config = typicalReplyConfig?.ConfigList?.find((_) => (_.Culture ?? null) === lang);
    }
    if (!config) {
      config = typicalReplyConfig?.ConfigList?.[0];
    }
    return config;
  }

  static async loadConfigForCurrentLanguageAndButtonId(culture, id) {
    const configForLang = await ConfigLoader.loadConfigForCurrentLanguage(culture);
    if (configForLang && configForLang.ButtonConfigList) {
      return configForLang.ButtonConfigList.find(conf => conf.Id === id);
    }
    return null;
  }
}
