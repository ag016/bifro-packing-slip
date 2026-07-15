# Kodus Home - Packing Slip Web App (Refactored)

A modern, responsive, and robust web application for creating packing slips and delivery notes. All data is saved to a Google Sheet, and slips can be downloaded as professional PDF documents.

## Reorganized Folder Structure

The project has been organized into logical directories for clean file management:
* **`src/`**: Holds all the source files that make up the Google Apps Script project.
  * `Code.gs` – Backend Google Apps Script (handles spreadsheet integration, Locks, validation, API).
  * `index.html` – Main application frame with Sidebar (Desktop) and Bottom Nav (Mobile).
  * `Styles.html` – Redesigned stylesheet implementing the slate-and-gold design system.
  * `Scripts.html` – Core router, auth, and dashboard controllers.
  * `ClientController.html` – Client directory actions (registration, display).
  * `SlipController.html` – Packing slip editor, visual print-preview, and filters.
  * `PdfGenerator.html` – High-quality PDF rendering engine using jsPDF.
* **`assets/`**: Static resource files (not uploaded to Apps Script).
  * `Kodus Logo.jpg` – High-res company logo image.
  * `Packing Slip - Google Sheets.pdf` – Example output PDF layout.

---

## Deployment Instructions

### Option A: Manual Copy-Paste (Recommended for Beginners)

1. **Create a Google Sheet**:
   * Open [Google Sheets](https://sheets.new) and create a blank spreadsheet (e.g. **Kodus Home - Packing Slips**).
2. **Open Apps Script**:
   * In the spreadsheet menu, go to **Extensions → Apps Script**.
   * Delete the default code inside `Code.gs`.
3. **Copy files from `src/`**:
   * Create the corresponding files inside the Apps Script online editor (for `.gs` select Script, for `.html` select HTML). Name them exactly as in the `src/` folder:
     1. `Code` (Script file) -> Copy from [Code.gs](file:///Users/akhilgupta/Desktop/Kodus/Packing%20Slip%20Software/src/Code.gs)
     2. `index` (HTML file) -> Copy from [index.html](file:///Users/akhilgupta/Desktop/Kodus/Packing%20Slip%20Software/src/index.html)
     3. `Styles` (HTML file) -> Copy from [Styles.html](file:///Users/akhilgupta/Desktop/Kodus/Packing%20Slip%20Software/src/Styles.html)
     4. `Scripts` (HTML file) -> Copy from [Scripts.html](file:///Users/akhilgupta/Desktop/Kodus/Packing%20Slip%20Software/src/Scripts.html)
     5. `ClientController` (HTML file) -> Copy from [ClientController.html](file:///Users/akhilgupta/Desktop/Kodus/Packing%20Slip%20Software/src/ClientController.html)
     6. `SlipController` (HTML file) -> Copy from [SlipController.html](file:///Users/akhilgupta/Desktop/Kodus/Packing%20Slip%20Software/src/SlipController.html)
     7. `PdfGenerator` (HTML file) -> Copy from [PdfGenerator.html](file:///Users/akhilgupta/Desktop/Kodus/Packing%20Slip%20Software/src/PdfGenerator.html)
4. **Authorize & Deploy**:
   * Save the project (click the Floppy disk icon).
   * Run the `doGet` function once from the toolbar to authorize spreadsheet access.
   * Click **Deploy → New deployment**.
   * Choose **Web app**, configure **Execute as: User accessing the web app** and **Who has access: Anyone with a Google account**.
   * Click Deploy and copy the Web App URL.

### Option B: Deploying with `clasp` (Advanced Developers)

If you are using Google's CLI tool `clasp`, configure your local `.clasp.json` in the root of the project to reference the `src` folder:

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "./src"
}
```

Then simply push using:
```bash
clasp push
```

---

## Technical Robustness Enhancements

1. **ID Collision Prevention**: Inside `Code.gs`, generated Client IDs and Packing Slip Codes are gated using Google's `LockService.getScriptLock()` with a 15-second timeout. This prevents duplicate IDs if multiple dispatchers click "Save" at the same instant.
2. **Server-Side Validation**: Sanitizes data on the Google Sheets backend to prevent empty item lists, invalid dates, negative quantities, or missing client names.
3. **Dynamic Delivery Card in PDF**: Automatically sizes the shaded client delivery panel in the output PDF to adapt to the length of the shipping address, eliminating clipping and text-overlaps.
4. **In-App Print Preview**: Before downloading, the view modal displays an exact visual facsimile of the printed packing slip, allowing checking of columns, addresses, and details on screen.
