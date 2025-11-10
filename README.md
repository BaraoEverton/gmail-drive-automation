# Gmail → Drive → Sheets → OCR Automation  
_Automated workflow using Google Apps Script (JavaScript)_

---

## 🧭 Overview (English)

This project automates the process of **importing Gmail attachments**, **saving them to Google Drive**, **logging metadata in Google Sheets**, and **extracting key data from PDFs using OCR**.  
It’s built entirely with **Google Apps Script (JavaScript)**, focusing on automation, reliability, and modular design for scalable document workflows.

---

## 🚀 Purpose

To centralize and process documents received via Gmail, maintaining structured logs in Google Sheets while automatically extracting essential data (e.g., invoice values, due dates, and document numbers) through OCR conversion.

---

## ⚙️ Core Modules

### 1. Gmail → Drive → Sheets  
**File:** `main.js`

Automates the import of Gmail attachments, uploads them to a Drive folder, and logs their metadata in Sheets.  
- Filters messages by **sender, domain, or subject keyword**  
- Downloads attachments and saves them in a predefined Drive folder  
- Logs in Sheets:
  - Message ID  
  - Date of email  
  - File name & Drive link  
  - Sender & recipient  
  - Processing status  
- Avoids duplicates using message IDs  
- Supports time-driven triggers for periodic execution  

---

### 2. Drive → OCR → Flow Dispatch  
**File:** `pdfExtractor.js`

Processes PDF files stored in Drive, converts them to Google Docs via OCR, extracts key data, and dispatches structured information to another spreadsheet (for example, an operations or billing flow).  
- Converts PDFs to text using OCR (via Advanced Drive API)  
- Extracts:
  - Value (`Valor`)
  - Due date (`Vencimento`)
  - Barcode / Boleto number  
  - Document number  
- Updates the same row in the original Sheet (columns H:L)  
- Appends the structured result to an external "Flow" spreadsheet  
- Automatically deletes temporary OCR documents  

---

## 🧩 Project Structure

