<div align="center">
  
  # 🤖 AC3110 DVT ISO Test App
  
  **Google Apps Script web app for reading and updating ISO test results stored in Google Sheets.**
  
  <p>
    <img src="https://img.shields.io/badge/Apps_Script-4285F4?style=for-the-badge&logo=google&logoColor=white" alt="Google Apps Script" />
    <img src="https://img.shields.io/badge/Google_Sheets-34A853?style=for-the-badge&logo=googlesheets&logoColor=white" alt="Google Sheets" />
    <img src="https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black" alt="JavaScript" />
    <img src="https://img.shields.io/badge/Clasp-000000?style=for-the-badge&logo=npm&logoColor=white" alt="Clasp" />
  </p>
</div>

---

## 🗺️ Architecture & File Map


<div align="center">
  <img width="1851" height="921" alt="image" src="https://github.com/user-attachments/assets/21ad570b-8b20-45ff-9881-e5cbe99402f9" />
</div>

<br>

### 🧩 Application Structure

| Category | Files | Description |
| :--- | :--- | :--- |
| ** پو Application Shell** | 📄 `server/Code.js`<br>📄 `server/AppCore.js`<br>📄 `index.html`<br>📄 `styles/CSS.html` | Web-app entry point, shared config, core namespaced layers, page layout, and shared styles. |
| **💾 Server-Side Data** | 📄 `server/Modules.js`<br>📄 `server/*_Config.js` | Generic ISO sheet discovery, normalization, generic CRUD, and feature-specific server services (e.g., `Lidar.js`, `WiFi.js`). |
| **🔌 Module Endpoints** | 📄 `HeadDepth.js`, `BaseDepth.js`<br>📄 `IRCliff.js`, `AprilTag.js`, etc. | Lightweight adapters exposing stable Apps Script function names while delegating work to `Modules.js`. |
| **🖥️ Browser-Side Partials** | 📄 `client/App_JS.html`<br>📄 `client/Modules_UI.html`<br>📄 `client/Overview_JS.html`<br>📁 `client/modules/*` | Browser registry, app state, navigation, overview matrix rendering, and feature-specific metadata/UI boundaries. |

---

## 📐 Development Conventions

To maintain a clean and scalable architecture, please adhere to the following rules:

1. **Global Settings:** Keep deployment-wide settings isolated in `Code.js`. Limit `Code.js` to public entry points and adapters.
2. **Shared Logic vs. Adapters:** Put shared server logic in `Modules.js`; keep individual module files as thin adapters.
3. **Data Encapsulation:** A feature owns its data access and renderer. Do not put LiDAR logic in `server/Modules.js`, or module-specific DOM code in `client/App_JS.html`. Sheet-specific behavior gets its own server file.
4. **Browser Boundaries:** Shared state/navigation lives in `client/App_JS.html`, overview-only rendering in `client/Overview_JS.html`, and each feature's form/table behavior in its exact matching `*_JS.html` partial.
5. **Naming & Namespaces:** 
   - Use a trailing underscore for private server helpers (e.g., `findIsoModuleSheet_`).
   - Prefer `ISOApp.*` namespaces and classes for new cross-cutting server code.
6. **Service APIs:** When adding a new service, expose a strict, small class API (`list`, `get`, `save`, `remove`) and keep all Google Sheets API calls contained within that service.

> **Note:** Apps Script loads top-level project files globally. For that reason, this project relies on clear file naming conventions and a strict file map rather than deeply nested folder trees.

---

## 🚀 GitHub & Apps Script Deployment

> **Production Script ID:** `1f4NuoWEDFo_fjguv6-N06MWr9prhbqVE52xTX0jiC4TJmIXnLJ7AGTgM`
> 
> ⚠️ **SECURITY WARNING:** This source connects to a writable Google Sheet. **Never commit OAuth tokens or passwords.** The registered tester directory is securely stored in Apps Script `ScriptProperties`, not in GitHub.

### 🌿 Branch-Based Testing Workflow

To safely develop new features without breaking production:

```bash
# 1. Sync local repository
git checkout main
git pull

# 2. Create a feature branch
git checkout -b feature/dvt3-robot

# 3. Develop & push to Apps Script
clasp push

# 4. Commit and push to GitHub
git add .
git commit -m "feat: add DVT3 robot support"
git push -u origin feature/dvt3-robot
