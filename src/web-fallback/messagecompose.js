/*
This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at http://mozilla.org/MPL/2.0/.

Copyright (c) 2026 ClearCode Inc.
*/
Office.onReady(() => {});

async function onNewMessageComposeCreated(event) {
  event.completed();
}
window.onNewMessageComposeCreated = onNewMessageComposeCreated;

Office.actions.associate("onNewMessageComposeCreated", onNewMessageComposeCreated);
