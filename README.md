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


---

## ⚙️ Setup

1. Open [Google Apps Script](https://script.google.com) and create a new project.  
2. Copy the `.js` files into the editor.  
3. Update configuration constants (`CONFIG`) in each file:
   - `GMAIL_SEARCH_QUERY`
   - `SUBJECT_KEYWORD`
   - `SHEET_NAME`
   - `DRIVE_FOLDER_ID`
   - External spreadsheet ID (for flow dispatch)
4. Enable the **Advanced Drive API** in the Apps Script editor.  
5. Run the main function (`iteraThreads()` or `extrairInfoPdf()`).  
6. Optionally, create a **time-driven trigger** for automation.  

---

## 🧠 Key Concepts

- Google Workspace Automation  
- Gmail + Drive + Sheets API Integration  
- OCR text extraction (Advanced Drive API)  
- Duplicate control with `Set()`  
- Logging and modular coding practices  

---

## 🧾 License
MIT License — free for personal or commercial use.

---

## 👤 Author
**Éverton Araújo**  
Automation · Google Workspace · Data & Workflow Engineering  
[LinkedIn](https://linkedin.com/in/seuusuario) • [GitHub](https://github.com/seuusuario)

---

---

# Gmail → Drive → Sheets → OCR Automação  
_Fluxo automatizado com Google Apps Script (JavaScript)_

---

## 🧭 Visão Geral (Português)

Este projeto automatiza o processo de **importar anexos do Gmail**, **salvar no Google Drive**, **registrar metadados no Google Sheets** e **extrair informações de PDFs via OCR**.  
Todo o fluxo é desenvolvido em **Google Apps Script (JavaScript)**, com foco em automação, rastreabilidade e modularidade.

---

## 🚀 Objetivo

Centralizar documentos recebidos por e-mail em uma pasta no Drive, mantendo um log estruturado na planilha e extraindo automaticamente dados essenciais (valor, vencimento, número do documento e código do boleto).

---

## ⚙️ Módulos Principais

### 1. Gmail → Drive → Sheets  
**Arquivo:** `main.js`

Automatiza a importação de anexos do Gmail, salvando-os no Drive e registrando os dados na planilha.  
- Filtro por **remetente, domínio ou palavra-chave**  
- Download automático dos anexos  
- Registro de dados na planilha:
  - ID da mensagem  
  - Data do e-mail  
  - Nome e link do arquivo  
  - Remetente e destinatário  
  - Status de processamento  
- Evita duplicidades (ID da mensagem)  
- Permite agendamento automático (gatilhos por tempo)  

---

### 2. Drive → OCR → Fluxos  
**Arquivo:** `pdfExtractor.js`

Processa os PDFs salvos no Drive, converte para Google Docs via OCR e extrai dados específicos.  
- Converte PDFs em texto com OCR (Drive API Avançada)  
- Extrai:
  - Valor  
  - Vencimento  
  - Código do boleto  
  - Número do documento  
- Atualiza automaticamente as colunas H:L na planilha  
- Envia os dados estruturados para outra planilha de fluxo  
- Remove os documentos temporários após a leitura  

---

## 🧩 Estrutura do Projeto

