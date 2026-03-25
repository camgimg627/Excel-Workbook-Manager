"use strict";

const appcontainer = require("office-addin-dev-settings/lib/appcontainer");
const originalEnsureLoopbackIsEnabled = appcontainer.ensureLoopbackIsEnabled;

appcontainer.ensureLoopbackIsEnabled = function ensureLoopbackIsEnabled(manifestPath) {
  return originalEnsureLoopbackIsEnabled(manifestPath, false);
};
