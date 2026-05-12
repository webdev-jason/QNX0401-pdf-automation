![Last Commit](https://img.shields.io/badge/last_commit-May-brightgreen) ![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=flat&logo=microsoft-excel&logoColor=white) ![VBA](https://img.shields.io/badge/VBA-1E3C6A?style=flat) ![Platform Windows](https://img.shields.io/badge/Platform-Windows-blue) ![License MIT](https://img.shields.io/badge/license-MIT-brightgreen)

# 📄 QNX0401 PDF Extraction & Download Macro

An Excel-based automation utility designed to extract serial numbers from scanned PDF documents and automatically fetch the corresponding test reports from a synced SharePoint directory. This macro utilizes invisible instances of Microsoft Word to process OCR data, transforming raw paper scans into a structured, actionable grid.

## ✨ Key Features

* **📄 Automated OCR Extraction:** Bypasses manual data entry by using Word's built-in conversion engine to read and extract grid tables directly from scanned `.pdf` files.
* **🧹 Data Sanitization & Formatting:** Auto-corrects common OCR typos (missing dashes, extra trailing characters), strictly filters for valid `JQ` serial number patterns, and flags errors with yellow highlighting for human review.
* **📊 UI/UX Optimization:** Dynamically spaces columns and applies zebra-striping to the extracted data to match the physical pages, making visual regression testing seamless.
* **☁️ SharePoint Syncing:** Iterates through the sanitized grid to search a local OneDrive/SharePoint directory (`GC_Outgoing_QC`) and automatically copies the latest `.pdf` test report for each valid serial number to the user's Downloads folder. 
* **🟢 Visual Status Indicators:** Updates the Excel UI in real-time, coloring cells Green for successful downloads and Red for files missing from the server.
