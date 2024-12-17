To set up and run the project that utilizes pdfplumber, openpyxl, pytesseract, and PIL to extract data from PDFs and save them into Excel sheets, follow these installation steps:

1. Clone the repository:
To clone the repository, run the following command:

   ``` bash
   git clone https://github.com/yourusername/your-repository.git
   cd your-repository


2. Create a virtual environment:
To isolate your project dependencies, create a virtual environment with the following command:

   ``` bash
   python -m venv .venv

3. Activate the virtual environment:
4. On Windows, run:

   ``` bash
   .venv\Scripts\activate

5. On macOS/Linux, run:

   ``` bash
   source .venv/bin/activate


6. Install required dependencies:
Install the necessary packages for the project by running:

   ``` bash
    pip install pdfplumber openpyxl pillow pytesseract

5. Install Tesseract OCR:
Since pytesseract is a Python wrapper for Tesseract OCR, you will need to install Tesseract OCR on your system.

6. On Windows:
Download the installer from the Tesseract GitHub Releases page.
Run the installer and follow the instructions.
Add the Tesseract executable to your system’s PATH, or set the pytesseract.pytesseract.tesseract_cmd variable to the Tesseract executable path in your script (e.g., C:\Program Files\Tesseract-OCR\tesseract.exe).

7. On macOS:
You can install Tesseract using Homebrew:

   ```bash
   brew install tesseract

8. On Linux (Ubuntu/Debian):

   ``` bash
   sudo apt update
   sudo apt install tesseract-ocr

9. Summary of Functions in the Script:
Extract Tables: Extracts tables from each page of the PDF using pdfplumber.
Write to Excel: Saves extracted tables to an Excel file using openpyxl. Each table is saved on a new sheet.
OCR Functionality: If needed, pytesseract can be used for OCR (optical character recognition) to read text from images within the PDF.
