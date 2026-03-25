"use strict";

const { spawnSync } = require("child_process");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");
const manifestPath = path.join(projectRoot, "manifest.xml");
const forwardedArgs = process.argv.slice(2);
const loopbackPatchPath = path.join(__dirname, "patch-office-loopback.js");

function resolveBin(binaryName) {
  return path.join(projectRoot, "node_modules", binaryName, "lib", "cli.js");
}

const cliCommand = process.execPath;
const debuggingCli = resolveBin("office-addin-debugging");
const devSettingsCli = resolveBin("office-addin-dev-settings");

function writeOutput(result) {
  if (result.stdout) {
    process.stdout.write(result.stdout);
  }
  if (result.stderr) {
    process.stderr.write(result.stderr);
  }
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    cwd: projectRoot,
    env: process.env,
    encoding: "utf8",
    stdio: options.capture ? ["inherit", "pipe", "pipe"] : "inherit",
  });

  if (options.capture) {
    writeOutput(result);
  }

  if (result.error) {
    throw result.error;
  }

  if ((result.status ?? 0) !== 0 && !options.allowFailure) {
    process.exit(result.status ?? 1);
  }

  return result;
}

function resolveAppName(args) {
  const defaultApp = (process.env.npm_package_config_app_to_debug || "excel").toLowerCase();
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if ((arg === "--app" || arg === "-a") && args[index + 1]) {
      return String(args[index + 1]).toLowerCase();
    }
  }
  return defaultApp;
}

function shouldSideload(args) {
  return !args.includes("--no-sideload");
}

function getProcessName(appName) {
  switch (appName) {
    case "excel":
      return "EXCEL";
    case "word":
      return "WINWORD";
    case "powerpoint":
      return "POWERPNT";
    case "outlook":
      return "OUTLOOK";
    default:
      return "";
  }
}

function getRunningWindowTitles(processName) {
  if (process.platform !== "win32" || !processName) {
    return [];
  }

  const command =
    `Get-Process ${processName} -ErrorAction SilentlyContinue ` +
    "| Select-Object -ExpandProperty MainWindowTitle";
  const result = spawnSync(
    "powershell",
    ["-NoProfile", "-Command", command],
    {
      cwd: projectRoot,
      env: process.env,
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
    }
  );

  if (result.error || (result.status ?? 0) !== 0) {
    return [];
  }

  return String(result.stdout || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function failIfOfficeAppIsOpen() {
  const appName = resolveAppName(forwardedArgs);
  const processName = getProcessName(appName);
  const windows = getRunningWindowTitles(processName);
  const displayName = appName.charAt(0).toUpperCase() + appName.slice(1);

  if (!windows.length) {
    return;
  }

  console.error(`Close all ${displayName} windows before running npm start.`);
  console.error("Open windows:");
  windows.forEach((title) => console.error(`  - ${title}`));
  console.error("");
  console.error(
    "Office sideloading reuses the open desktop session. If Excel is already holding a workbook"
  );
  console.error(
    "that contains an older development add-in, Office shows the 'add-in is no longer available'"
  );
  console.error("message instead of loading this project cleanly.");
  process.exit(1);
}

function ensureLoopback() {
  if (process.platform !== "win32") {
    return;
  }

  console.log("Ensuring localhost loopback is enabled for the Office webview...");
  run(
    cliCommand,
    [devSettingsCli, "appcontainer", "edgewebview", "--loopback", "--yes"],
    { allowFailure: true }
  );
}

function clearStaleDebugState() {
  console.log("Clearing any stale add-in debugging state...");
  run(cliCommand, [debuggingCli, "stop", manifestPath], { allowFailure: true });
}

function startDebugging() {
  console.log("Starting the local Excel add-in session...");
  const result = run(
    cliCommand,
    ["-r", loopbackPatchPath, debuggingCli, "start", manifestPath, ...forwardedArgs],
    { allowFailure: true }
  );
  process.exit(result.status ?? 0);
}

if (shouldSideload(forwardedArgs)) {
  failIfOfficeAppIsOpen();
}
ensureLoopback();
clearStaleDebugState();
startDebugging();
