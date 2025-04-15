# pdf_watcher
A Windows tool that watches folders, OCR-scans PDFs, detects labels, and auto-prints via SumatraPDF. Runs from the system tray with Start/Stop, logs activity, and bundles Tesseract for easy setup.


This Windows utility monitors multiple folders for newly added PDF files. It automatically processes each PDF using built-in text extraction and OCR (via Tesseract), identifies the document type based on content or barcode IDs, and sends it to the correct printer (A4 or A6) using SumatraPDF for silent printing.

Key Features
Auto-Printing: Detects new PDFs and sends them to the appropriate printer automatically.

OCR Support: Uses Tesseract to recognize text or barcodes from image-based or scanned PDFs.

Multiple Folder Watchers: Watches customizable directories and assigns different logic per folder.

Smart Routing: Matches text or ID patterns to known couriers (e.g., DHL, ELTA) and prints accordingly.

System Tray App: Runs quietly in the background with a tray icon for Start, Stop, and Quit controls.

Logging: Keeps a log file with timestamps and detailed processing output, including errors.

Bundled Installer: Comes with Tesseract setup and a portable SumatraPDF, so it works out of the box.

Usage
Just run the executable and control it via the system tray. PDF detection and printing happens silently in the background. You can customize watched folders, printers, and detection patterns via the source code.
