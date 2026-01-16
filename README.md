# EU VAT Number Validator for Google Sheets

Validate EU VAT numbers directly in **Google Sheets** using the official European Commission **VIES (VAT Information Exchange System)** service.

This project provides a custom Google Apps Script function you can paste into a Sheet and use like any other formula—great for invoicing, accounting workflows, and EU cross-border compliance checks.

---

## ✨ Features

- ✅ Validates EU VAT numbers using the official VIES database  
- 📊 Works natively inside Google Sheets (Google Apps Script)  
- 🔄 Supports validating many VAT numbers by filling formulas down  
- 🧩 Returns structured results (date, valid/invalid, name, address, error)  
- ⚡ Lightweight and easy to install

---

## 🔗 References

- **WSDL:** `https://ec.europa.eu/taxation_customs/vies/services/checkVatService.wsdl`  
- **VIES FAQ:** `https://ec.europa.eu/taxation_customs/vies/faq.html`

---

## 📦 Installation Guide

### Step 1: Open your Google Sheet
Open the Google Sheet where you want to validate VAT numbers.

### Step 2: Access Apps Script Editor
- Click **Extensions** in the menu bar  
- Select **Apps Script**  
- The Apps Script editor opens in a new tab

### Step 3: Create / prepare a script file
- In the editor you’ll see a default `Code.gs`
- Delete any existing code in `Code.gs` *(or create a new file)*
- To create a new file: click the **+** next to **Files** → **Script**
- Name it something like `CheckVAT` *(optional)*

### Step 4: Copy the script
Copy the JavaScript from this repository and paste it into the Apps Script editor.

### Step 5: Save the project
- Click the disk icon or press **Ctrl+S** (**Cmd+S** on Mac)
- Give the project a name like **VAT Validator**

### Step 6: Authorization (first run only)
The first time you use the function, Google will ask for permissions:
- Click **Review permissions**
- Select your Google account
- If you see a warning, click **Advanced**
- Click **Go to [Project Name] (unsafe)** *(it’s your own script)*
- Click **Allow**

---

## 🧪 Usage in Google Sheets

### Basic usage

In any cell, type:

```text
=CHECK_VAT_SERVICE("NL123456789B01")
