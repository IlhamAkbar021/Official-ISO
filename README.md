# AC3110 DVT ISO Test App

Google Apps Script web app for reading and updating ISO test results stored in
Google Sheets.


## File map
<img width="1902" height="954" alt="image" src="https://github.com/user-attachments/assets/da0949cd-e80c-4765-b468-a1c6629fb6d4" />


### Application shell

- `server/Code.js` — web-app entry point, shared configuration, and HTML includes.
- `server/AppCore.js` — namespaced core layer: configuration, tester service, and web-app controller.
- `index.html` — page layout and HTML partial includes.
- `styles/CSS.html` — shared styles.

### Server-side data access

- `server/Modules.js` — generic ISO sheet discovery, normalization, image loading, and
  generic CRUD operations.
- `server/ISOConfig.js`, `server/Overview.js`, `server/StaticISO.js`,
  `server/Lidar.js`, and `server/WiFi.js` — feature-specific server services.

### Module endpoint files

`HeadDepth.js`, `BaseDepth.js`, `IRCliff.js`, `RatedSpeed.js`, `EStop.js`,
`TurningWidth.js`, `CliffLowerSpeed.js`, `TraverseGap.js`, `NarrowCorridor.js`,
and `AprilTag.js` are intentionally small. They expose stable Apps Script
function names while delegating shared work to `Modules.js`.

### Browser-side feature partials

- `client/App_JS.html` — application state, navigation, forms, and shared actions.
- `client/Modules_UI.html` — browser registry for all module identities and renderers.
- `client/modules/static/BaseDepth_JS.html` — Base Depth Camera browser metadata and UI boundary.
- `client/modules/static/IRCliff_JS.html` — IR Cliff browser metadata and UI boundary.
- `RatedSpeed_JS.html`, `EStop_JS.html`, `TurningWidth_JS.html`,
  `CliffLowerSpeed_JS.html`, `TraverseGap_JS.html`, `NarrowCorridor_JS.html`,
  and `AprilTag_JS.html` — Dynamic ISO browser metadata and UI boundaries.
- `client/Overview_JS.html` — overview matrix rendering and editing.
- `client/modules/static/Lidar_JS.html` — LiDAR form and rendering logic.
- `client/modules/static/WiFi_JS.html` — Wi-Fi form and rendering logic.

## Conventions

1. Keep deployment-wide settings in `Code.js`.
2. Put shared server logic in `Modules.js`; keep module files as adapters.
3. Put sheet-specific behavior in its own server file.
4. Keep browser code in the matching `*_JS.html` partial.
5. Use a trailing underscore for private server helpers, such as
   `findIsoModuleSheet_`.

6. Prefer `ISOApp.*` namespaces and classes for new cross-cutting server code;
   keep `Code.js` limited to Apps Script public entry points and compatibility
   adapters.
7. A feature owns its data access and renderer: do not put LiDAR logic in
   `server/Modules.js`, or module-specific DOM code in `client/App_JS.html`.
8. Browser code should follow the same boundary: shared state/navigation in
   `client/App_JS.html`, overview-only rendering in `client/Overview_JS.html`, and each
   feature's form/table behavior in its matching partial.
9. When adding a service, expose a small class API (`list`, `get`, `save`,
   `remove`) and keep Google Sheets calls inside that service.

Apps Script loads top-level project files globally. For that reason, this
project uses clear file sections and a file map instead of nested folders.

## GitHub and Apps Script

Production Script ID: `1f4NuoWEDFo_fjguv6-N06MWr9prhbqVE52xTX0jiC4TJmIXnLJ7AGTgM`

The Script ID is an identifier, but this source connects to a writable Google Sheet. Never commit OAuth tokens or
passwords. The registered tester directory is stored in Apps Script
`ScriptProperties`, not in GitHub.

### Branch-based testing

```bash
git checkout main
git pull
git checkout -b feature/dvt3-robot

Update Code
clasp push
git add .
git commit -m "Add DVT3 robot support"
git push -u origin feature/dvt3-robot
```

Test the feature branch first, then merge it through a GitHub Pull Request.
For real isolation, the testing branch must use a separate Apps Script project
and test spreadsheet. `clasp push` updates the Script ID in `.clasp.json`.

### Another user continues the project

```bash
git clone https://github.com/IlhamAkbar021/Official-ISO.git
cd Official-ISO
npm install -g @google/clasp
clasp login (Company account)
clasp push
```

That user must have access to both the Apps Script project and the connected
Google Sheet.
