# AC3110 DVT ISO Test App

Google Apps Script web app for reading and updating ISO test results stored in
Google Sheets.

## File map

### Application shell

- `Code.js` — web-app entry point, shared configuration, and HTML includes.
- `index.html` — page layout and HTML partial includes.
- `CSS.html` — shared styles.

### Server-side data access

- `Modules.js` — generic ISO sheet discovery, normalization, image loading, and
  generic CRUD operations.
- `ISOConfig.js` — module names, aliases, phases, and sheet matching rules.
- `Overview.js` — overview matrix data and batch updates.
- `StaticISO.js` — shared sheet lookup helper for fixed-format static data.
- `Lidar.js` — LiDAR-specific split-table reader and CRUD operations.
- `WiFi.js` — Wi-Fi-specific reader and CRUD operations.

### Module endpoint files

`HeadDepth.js`, `BaseDepth.js`, `IRCliff.js`, `RatedSpeed.js`, `EStop.js`,
`TurningWidth.js`, `CliffLowerSpeed.js`, `TraverseGap.js`, `NarrowCorridor.js`,
and `AprilTag.js` are intentionally small. They expose stable Apps Script
function names while delegating shared work to `Modules.js`.

### Browser-side feature partials

- `App_JS.html` — application state, navigation, forms, and shared actions.
- `Overview_JS.html` — overview matrix rendering and editing.
- `Lidar_JS.html` — LiDAR form and rendering logic.
- `WiFi_JS.html` — Wi-Fi form and rendering logic.

## Conventions

1. Keep deployment-wide settings in `Code.js`.
2. Put shared server logic in `Modules.js`; keep module files as adapters.
3. Put sheet-specific behavior in its own server file.
4. Keep browser code in the matching `*_JS.html` partial.
5. Use a trailing underscore for private server helpers, such as
   `findIsoModuleSheet_`.

Apps Script loads top-level project files globally. For that reason, this
project uses clear file sections and a file map instead of nested folders.

## GitHub and Apps Script

Production Script ID: `1f4NuoWEDFo_fjguv6-N06MWr9prhbqVE52xTX0jiC4TJmIXnLJ7AGTgM`

Keep the GitHub repository private. The Script ID is an identifier, but this
source connects to a writable Google Sheet. Never commit OAuth tokens or
passwords. The registered tester directory is stored in Apps Script
`ScriptProperties`, not in GitHub.

### Branch-based testing

```bash
git checkout main
git pull
git checkout -b feature/dvt3-robot
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
git clone https://github.com/<USER_OR_ORG>/<REPOSITORY>.git
cd <REPOSITORY>
npm install -g @google/clasp
clasp login
clasp push
```

That user must have access to both the Apps Script project and the connected
Google Sheet.
