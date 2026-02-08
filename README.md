### hr-travel-docs-automation
### HR Operations
### Visa Travel Letter Automation

### Overview

Python automation to generate **Visa Entry and Interview letters** for employees using Excel data and Word Mail Merge templates.

Built to eliminate manual document creation for business travel and visa processing.

---

### What it does

* Reads employee travel data from Excel (.xlsm)
* Filters valid records
* Selects template based on:

  * Letter Type (Entry / Interview)
  * Travel Location (Kansas City / Malvern)
* Auto-populates Word Mail Merge fields:

  * Employee details
  * Travel dates
  * Job information
  * Stay address, passport details
* Applies gender-based pronouns automatically
* Generates individual DOCX letters named by Employee/Operator ID

---

### Tech Stack

Python, pandas, xlwings, MailMerge

---

### How to run

1. Update:

   * Network folder paths
   * Template locations
   * Excel file path
2. Ensure required templates and data file are available
3. Run:

```bash
python script.py
```

---

### Notes

* Designed for internal HR visa/travel processing
* Requires shared drive access and standardized templates
* Excel column names must match expected fields

---

**Year:** ~2018–2020
Reference project demonstrating bulk document generation and HR process automation.
