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
```

## Input format

- The VAT number **must include the 2-letter EU country code** (ISO 3166-1 alpha-2), e.g. `NL`, `BE`, `DE`, `FR`
- The country code must be **followed directly by the VAT number**
- **Spaces, dots, and other special characters are automatically removed**
- Input is **case-insensitive**

### Examples

```text
NL123456789B01
BE0123456789
DE123456789
FR12345678901
```

## Output format

The function returns an array that expands into **5 columns** in Google Sheets:

```text
[Request Date, Valid, Company Name, Address, Error Message]
```

### Column details

| Column | Name           | Description                                                                 |
|--------|----------------|-----------------------------------------------------------------------------|
| A      | Request Date   | Date and time of the VIES request (e.g. `2015-03-09+01:00`)                  |
| B      | Valid          | Boolean result: `true` if the VAT number is valid, otherwise `false`        |
| C      | Company Name   | Registered company name returned by VIES (may be empty)                    |
| D      | Address        | Registered address returned by VIES (may contain line breaks)              |
| E      | Error Message  | Empty if successful, otherwise an error or status message                  |

## Example spreadsheet layout

### Input and output in Google Sheets

| Column A     | Column B                     | Column C       | Column D | Column E        | Column F        |
|--------------|------------------------------|----------------|----------|-----------------|-----------------|
| VAT Number   | Formula                      | Request Date   | Valid    | Company Name    | Address         |
| NL1234567…   | =CHECK_VAT_SERVICE(A2)        | 2015-03-09…    | true     | ACME CORP B.V.  | KERKSTRAAT…     |
| BE0123456…   | =CHECK_VAT_SERVICE(A3)        | 2015-03-09…    | false    |                 | INVALID_INPUT   |

### Notes

- Enter VAT numbers in **Column A**
- Place the formula in **Column B**
- The result automatically expands across **5 columns** (B → F)
- Copy the formula down to validate multiple VAT numbers

## Advanced usage — multiple VAT numbers
## Handling errors

If a validation fails, the **Error Message** column will contain a descriptive message.

Common cases include:

- **Network errors**  
  `Network error: [message]`

- **Invalid input format**  
  `Invalid VAT number: too short`

- **Service errors**  
  `HTTP error 500: Service unavailable`

- **Invalid VAT number**  
  VIES may return codes such as `INVALID_INPUT` or similar

If many rows fail at the same time, the VIES service may be temporarily unavailable.

---

## Rate limiting

To prevent overloading the EU VIES service, the script enforces a **1-second delay** between requests.

For larger batches, expect approximately:

- **~60 VAT validations per minute**

---

## Notes & tips

- The VIES service can be temporarily unavailable due to maintenance
- Not all EU VAT numbers appear in VIES immediately after registration
- Some member states return limited data (company name and/or address may be empty)
- To avoid repeated requests, consider copying the results and **pasting values**
- Saving validated results improves performance and reliability

1. Put VAT numbers in **column A** (starting at row 2)
2. In **B2**, enter:
```text
=CHECK_VAT_SERVICE(A2)
```
3. The result will automatically expand across **5 columns** (B → F)
4. Copy/fill the formula down to validate multiple VAT numbers

