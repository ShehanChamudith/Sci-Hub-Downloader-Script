# 📄 Sci-Hub DOI PDF Downloader with Status Tracker

This Python script automates the process of retrieving academic paper PDFs from [Sci-Hub](https://sci-hub.se) using DOIs listed in a CSV file. It simulates human-like interactions to avoid bot detection, downloads valid PDFs, and updates the original dataset with download status.

---

## ✅ Features

- 🔍 Searches Sci-Hub using DOIs from a CSV
- 👨‍💻 Simulates human typing via Selenium Actions
- ⏳ Introduces random delays between actions (to mimic human behavior)
- 📥 Downloads and validates PDF files
- 📊 Updates an Excel file with download status:
  - ✅ Green for "Found"
  - ⚠️ Orange for "Not Found"
- 💾 Saves invalid or broken content responses as `.html` for debugging

---

## 📁 Input CSV Format

The input CSV file must include at least the following columns:

- `DOI`
- `Title`

Example:

| DOI                       | Title                                          |
|--------------------------|------------------------------------------------|
| 10.3390/su15054017       | Attributes of Diffusion...                    |
| 10.1186/s40691-017-0119-8| Understanding millennial consumer’s adoption...|

---

## 🛠 Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/scihub-downloader.git
   cd scihub-downloader

2. **Install Dependencies**:
    ```bash
    pip install pandas tqdm requests PyPDF2 openpyxl selenium undetected-chromedriver

3. **Run the Script**:
    ```bash
    python script.py

## ⚠️ Disclaimer

-This script accesses Sci-Hub, which may be restricted in your country or institution.
-This project is intended strictly for academic, educational, and research use only.
-Use responsibly and comply with local laws.
