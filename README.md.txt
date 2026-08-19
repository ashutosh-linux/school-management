# 🏫 School ERP & Smart ID Card Studio

A comprehensive, offline-first desktop application engineered to streamline daily operations for educational institutes. Built entirely in Python, this system manages student admissions, fee ledgers, and financial tracking, while featuring a highly advanced, algorithmic PDF generation engine for CR80 ID cards and admission forms.

![Python](https://img.shields.io/badge/Python-3.x-blue?style=for-the-badge&logo=python)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-darkgreen?style=for-the-badge)
![SQLite](https://img.shields.io/badge/Database-SQLite-lightgrey?style=for-the-badge&logo=sqlite)
![ReportLab](https://img.shields.io/badge/PDF-ReportLab-red?style=for-the-badge)

## ✨ Key Features

*   **📊 Financial Dashboard:** Real-time metrics on total students, billed amounts, collected revenue, outstanding dues, and net income.
*   **👨‍🎓 Advanced Admission Desk:** Complete student lifecycle management. Captures 36+ unique data points and auto-generates official, printable PDF Admission Forms.
*   **💳 Billing & Payments Hub:** 
    *   Generate individual PDF Demand Bills with itemized fee breakdowns.
    *   Track partial/full payments and auto-generate PDF Receipts.
    *   One-click WhatsApp integration to send fee reminders directly to parents.
*   **🪪 Smart ID Card Studio:** 
    *   Batch generates standard CR80 ID cards formatted flawlessly on A4 grids (10 per page).
    *   Utilizes mathematical clipping paths for custom wavy UI designs.
    *   Auto-injects dynamic student QR codes.
    *   Smart background removal (RGB masking) for transparent Principal Signature stamping.
*   **📥 Excel Bulk Operations:** Seamlessly import hundreds of students at once using `.xlsx` templates.
*   **💾 Secure & Local:** Runs entirely offline using SQLite, with automated localized backups ensuring maximum data privacy and zero reliance on cloud subscriptions.

## 📸 Screenshots

*(Add screenshots of your application here by placing images in an `assets` folder!)*
> `![Dashboard](assets/dashboard.png)`
> `![ID Card Generator](assets/id_cards.png)`
> `![Admission Form](assets/admission.png)`

## 🛠️ Tech Stack & Libraries

*   **Core:** Python 3.x
*   **UI Interface:** `tkinter`, `ttk` (Custom dark-mode UI styling)
*   **Database:** `sqlite3`
*   **PDF Generation:** `reportlab` (Canvas, Shapes, Barcode/QR, Tables, Paragraphs)
*   **Data Handling:** `openpyxl`
*   **Packaging & Deployment:** `PyInstaller` (Executable generation) & `Inno Setup` (Windows Installer)

## 🚀 Running the Project Locally

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/ashutosh-linux/school-management.git](https://github.com/ashutosh-linux/school-management.git)
   cd school-management

Install required dependencies:

pip install reportlab openpyxl

Run the application:

python fees_app.py


📦 Building the Executable (.exe)
To compile this project into a standalone Windows executable with the bundled icon and ReportLab dependencies:
python -m PyInstaller --noconsole --onefile --clean --collect-all reportlab --add-data "school_logo.ico;." -i school_logo.ico fees_app.py



## 👨‍💻 Developer
Developed by **Ashutosh Kumar** 
Connect with me on [LinkedIn](https://www.linkedin.com/in/ashutosh-kumar625490) | [GitHub](https://github.com/ashutosh-linux)