# Server architecture

The Apps Script runtime places every `.js` file in one global scope. This
project therefore uses an explicit boundary between public RPC functions and
private implementation code.

## Layers

1. `AppCore.js` — application configuration and web controller.
2. `ServerLibrary.js` — reusable server infrastructure (`Result`, `Cache`,
   `Database`, `Rows`, and `Http`). Its IIFE keeps implementation helpers
   private.
3. `ISOConfig.js` — module registry and sheet aliases.
4. Domain files (`Lidar.js`, `WiFi.js`, `HeadDepth.js`, etc.) — business rules
   exposed through `ISOApp.Services.<Domain>`.
5. Global adapter functions — the only functions intended for
   `google.script.run`, triggers, template scriptlets, or web-app entry points.

## Visibility convention

- **Public RPC:** a small global function that delegates immediately to a
  service, for example `getWifiData()`.
- **Private helper:** enclosed in an IIFE when possible; otherwise its name ends
  in `_`, for example `normalizeLidarImportPayload_()`.
- **Shared internal API:** namespaced under `ISOApp.Server`; do not add new
  generic global helpers.

## Adding a server feature

1. Put reusable infrastructure in `ServerLibrary.js`.
2. Put feature rules in `ISOApp.Services.<Domain>`.
3. Add the smallest possible global RPC adapter required by the browser.
4. Return `{ success: true, ... }` or `{ success: false, error: "..." }` using
   `ISOApp.Server.Result`.
5. Access the workbook through `ISOApp.Server.Database` and cache through
   `ISOApp.Server.Cache`.

This keeps frontend contracts stable while preventing unrelated server modules
from depending on each other's implementation details.
