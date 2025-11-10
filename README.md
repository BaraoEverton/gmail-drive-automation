# gmail-drive-automation
Automation script for importing Gmail attachments into Google Drive and Sheets (Apps Script)

# Gmail → Drive → Sheets Automation  
_Automated workflow using Google Apps Script (JavaScript)_

---

## 🧭 Overview (English)

This project automates the process of **extracting Gmail attachments**, saving them to **Google Drive**, and logging key metadata into a **Google Sheets** file.  
It’s built with **Google Apps Script**, designed for efficiency, transparency, and reuse in different administrative contexts.

### 🚀 Purpose
Centralize incoming documents from Gmail in a single Drive folder while keeping a structured log in Sheets — avoiding duplicates and simplifying auditing.

### ⚙️ Features
- Automatically searches emails by **sender, domain, or keyword**
- Downloads attachments and uploads them to a chosen Drive folder
- Logs message metadata in Sheets:
  - Message ID  
  - Email date  
  - File name and Drive link  
  - Sender and recipient  
  - Receipt status  
- Prevents duplication using the message ID as a unique key
- Supports scheduled runs via **time-driven triggers**

### 🧩 Project Structure
