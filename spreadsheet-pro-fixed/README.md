# Spreadsheet Pro

An interactive, high-fidelity, client-side spreadsheet application featuring dynamic multi-sheet capabilities, advanced mathematical and logical formula calculations, responsive grid layout resizing, and a powerful **Document Data Import Center** supporting text, CSV, PDF, and Word documents.

---

## 🚀 Core Features

1. **Robust Formula Engine**:
   - Math Functions: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `ABS`, `SQRT`, `ROUND`
   - String Editing: `LEN`, `UPPER`, `LOWER`, `LEFT`, `RIGHT`, `CONCAT`
   - Evaluation Logic: `IF`, `AND`, `OR`, `NOT`
   - Lookups & Indexes: `VLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`
   - Conditional Aggregates: `SUMIF`, `COUNTIF`, `AVERAGEIF`
   - Utility Helpers: `TODAY()`, `NOW()`, `ISBLANK()`, `ISNUMBER()`

2. **Document Data Import Center**:
   - **CSV / TXT**: Seamless comma or tab delimiter layout mapping.
   - **PDF & Word (.docx/.doc)**: Standard client-side layout parser. Scans document text streams and parses them into a tabular structure with a live staging spreadsheet preview before writing to your active sheet.
   - **Autofill Selection**: Replicate formula calculations or linear sequences over selected boundaries instantly.

3. **Advanced Operations Toolbar**:
   - Undo/Redo historical revision states.
   - High-precision column resizing handles.
   - Multi-tab sheet adding, renaming, and secure deletion.
   - Single-cell and range selections.
   - Column sorting (A-Z, Z-A) and universal "Find & Replace" across active sheets.
   - Advanced text alignment, borders, and coloring (background fill & text color).

---

## 🛠️ File Structure

* `/src/types.ts`: Global application types and stylesheet interfaces.
* `/src/App.tsx`: Central spreadsheet state machine. Handles cell values, stylesheet properties, active selections, sheet changes, and undo/redo histories.
* `/src/components/Grid.tsx`: Core grid canvas element that renders column coordinates, row values, selection borders, and coordinates live editing.
* `/src/components/Toolbar.tsx`: Design styling operations, find & replace toggles, active sorting, cell background/text formats, and export controls.
* `/src/components/FormulaBar.tsx`: Dynamic cell formula edit node.
* `/src/components/ImportModal.tsx`: Visual Drag-and-Drop file processing center for CSV, PDF, Word, and TXT files.
* `/src/utils/formulaEvaluator.ts`: Interactive cell evaluator. Supports range parsing, mathematical orders of operation, cell reference dependencies, and robust helper evaluations.

---

## 💡 Navigation Cheat Sheet

- **Typing Data**: Double-click any cell or highlight it and start typing. Press `Enter` to scroll down or `Tab` to scroll right.
- **Formulas**: Precede any formula with an equals sign (e.g., `=SUM(B5:B10)`). Matches are live!
- **Auto-Fill Selection**: Highlight a range starting with a populated cell and click the "Auto-Fill Selection (Ctrl+D)" button on the toolbar or press `Ctrl+D` to propagate values down.
- **Help Panel**: Hover or click the **Help Tips** button in the header at any time to inspect formulas and hotkeys.

---

## 🚀 How to Deploy

You can deploy and share this fully-functional application using multiple methods:

### 1. Instant Deployment via Google AI Studio Build
Google AI Studio Build provides standard, native deployment tools directly within your current workspace environment:
* **One-Click Share/Deploy**: Look at the top right of the Google AI Studio interface. Click the **Share** or **Deploy** button to provision a fully-optimized public link. This will compile your React applet into static assets and deploy them onto Google Cloud Run.
* **Adding/Managing Secrets**: If you introduce backend APIs or services in the future, you can configure environmental secrets securely via the **Settings (Gear Icon)** panel in the top-right corner.

### 2. Exporting the Workspace Code
If you want to move your codebase outside AI Studio, utilize these export pathways under the **Settings/Menu** option:
* **Export to GitHub**: Connect your GitHub account to automatically push this exact repository, establishing permanent revision control.
* **Download ZIP**: Retrieve a clean, fully configured `.zip` file of the workspace to inspect or move to another hosting platform (Vercel, Netlify, AWS, etc.).

### 3. Local Development & Production Build
To spin up the server locally on your machine, follow these commands in your standard terminal:

1. **Install Dependencies**:
   ```bash
   npm install
   ```
2. **Launch the Development Server**:
   ```bash
   npm run dev
   ```
3. **Compile for Production**:
   ```bash
   npm run build
   ```
   This compiles clean, highly optimized static production assets directly into the `/dist` bundle, ready to be served by any static file hosting service.

---

## 💻 How to Package as a Desktop App (.EXE, .DMG, or Linux App)

We have pre-configured **Electron** and **Electron Builder** inside this repository so you can easily compile this web app into a fully standalone desktop executable!

### Step 1: Open Terminal on Your Local PC
Navigate to your extracted folder containing these workspace files:
```bash
cd /path/to/spreadsheet-pro
```

### Step 2: Install Local Packages
Run the install command to fetch all normal web and native packaging packages (including Electron):
```bash
npm install
```

### Step 3: Run as a Desktop App (Development Mode)
To spin up the application inside a native Desktop wrapper window instead of a standard browser tab:
* First, launch the dev server:
  ```bash
  npm run dev
  ```
* Then, in a second terminal window, launch the Electron app:
  ```bash
  ELECTRON_DEV=true npm run electron:dev
  ```

### Step 4: Package into a Standalone Desktop Executable (.EXE / .DMG)
To compile and package your application into a fully-contained executable file that runs natively on your computer without needing any terminal commands or hosting:
```bash
npm run electron:build
```
This single command will:
1. Compile your optimized React + Tailwind production assets into the `/dist` directory.
2. Package the app bundle with the Chromium rendering engine inside a native desktop frame.
3. Output the standalone installer files (e.g., `.exe` setup on Windows, `.dmg` on macOS, or `.AppImage` on Linux) inside the **`/dist-desktop`** directory!

You can double-click this `.exe` file to install it directly onto your computer, running Spreadsheet Pro just like Microsoft Excel or Google Sheets natively, completely offline!


