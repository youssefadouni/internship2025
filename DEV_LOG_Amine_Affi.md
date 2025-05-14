# DYDON AI – Internship 2025 Application Submission

This repository contains my solution for the DYDON AI Summer Internship 2025 programming task, along with an enhanced version that demonstrates advanced capabilities.

---

## 📝 Task Overview

The goal was to implement a Python function to extract text from various file types (`.pdf`, `.docx`, `.xlsx`) located in an `uploads/` directory. A key requirement was to use Tesseract OCR as a fallback for PDF files that do not contain extractable text.

---

## ⚙️ My Approach

My solution, `extract_text.py`, implements the following:

1.  **File Iteration**: The script scans the `uploads/` directory for files.
2.  **Type-Specific Extraction**:
    *   **PDFs**: Uses `pdfplumber` to extract text. If no text is found, it converts PDF pages to images and uses `pytesseract` (with Pillow for image handling) to perform OCR.
    *   **DOCX**: Uses `python-docx` to read and extract text from `.docx` files.
    *   **XLSX**: Uses `openpyxl` to iterate through cells and extract their string content from `.xlsx` files.
3.  **Error Handling & Logging**: Implemented robust error handling for each file type and operation. All significant events, errors, and processing steps are logged to `extraction.log` and also printed to the console for immediate feedback.
4.  **Modularity**: The extraction logic for each file type is encapsulated in separate functions for clarity and maintainability (`extract_text_from_pdf`, `extract_text_from_docx`, `extract_text_from_xlsx`).
5.  **Main Script Logic**: The `process_files_in_uploads()` function orchestrates the file processing, and the `if __name__ == "__main__":` block handles the execution, including a check for Tesseract's availability.

---

## 🚀 Enhanced Version

I've also created an enhanced version of the text extraction tool (`enhanced_extract_text.py`) that demonstrates more advanced capabilities:

### Advanced Features

- **Multi-threading**: Parallel processing of files for improved performance
- **Progress Visualization**: Real-time progress bars during extraction
- **Extended Format Support**: Added support for TXT, CSV, and JSON files
- **Text Analysis**: Keyword extraction, language detection, and sentiment analysis
- **Improved Error Handling**: More robust error recovery and detailed reporting
- **Command-line Interface**: Flexible options for customizing extraction behavior

### Using the Enhanced Version

```bash
python enhanced_extract_text.py --analyze --verbose
```

Options:
- `-i, --input`: Specify input directory (default: uploads)
- `-o, --output`: Specify output directory (default: extracted_texts)
- `-w, --workers`: Number of parallel workers (default: 4)
- `-a, --analyze`: Perform text analysis (keywords, sentiment, etc.)
- `-v, --verbose`: Enable detailed logging

---

## 🛠️ Dependencies

The project requires Python 3.8+ and the following libraries. You can install them using the provided `requirements.txt` file:

```bash
pip install -r requirements.txt
```

### Basic Version Dependencies
-   `pytesseract`: For OCR capabilities.
-   `pdfplumber`: For PDF text extraction and page manipulation.
-   `python-docx`: For reading `.docx` files.
-   `openpyxl`: For reading `.xlsx` files.
-   `Pillow`: Required by `pytesseract` for image processing (especially for OCR from PDF pages).

### Enhanced Version Additional Dependencies
-   `tqdm`: For progress bars
-   `nltk`: For natural language processing and keyword extraction
-   `langdetect`: For language detection
-   `textblob`: For sentiment analysis

**Additionally, Tesseract OCR must be installed on your system.**

-   **On macOS (via Homebrew):**
    ```bash
    brew install tesseract
    brew install tesseract-lang # For language packs, if needed
    ```
-   **On Linux (Debian/Ubuntu):**
    ```bash
    sudo apt update
    sudo apt install tesseract-ocr
    sudo apt install tesseract-ocr-eng # For English language pack
    ```
-   **On Windows:** Download the installer from the [official Tesseract GitHub page](https://github.com/UB-Mannheim/tesseract/wiki) and ensure the installation directory is added to your system's PATH, or set `pytesseract.pytesseract.tesseract_cmd` in the script.

---

## 🚀 How to Run

1.  **Clone the repository** (if you haven't already).
2.  **Create a virtual environment** (recommended):
    ```bash
    python3 -m venv venv
    source venv/bin/activate  # On Windows: venv\Scripts\activate
    ```
3.  **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
4.  **Ensure Tesseract OCR is installed** and accessible (see Dependencies section).
5.  **Create an `uploads/` directory** in the root of the project if it doesn't exist.
6.  **Place your `.pdf`, `.docx`, and `.xlsx` files** into the `uploads/` directory.
7.  **Run the script**:
    ```bash
    python extract_text.py
    ```

The script will process the files, print a summary of extracted text (first 200 characters per file) to the console, and log detailed information to `extraction.log`.

---

## 🎯 Focus on Evaluation Criteria

-   **Code Quality and Structure**: The code is organized into functions with clear responsibilities. Logging is used for traceability.
-   **Elegance and Simplicity**: Standard libraries are used as requested. The OCR fallback for PDFs is integrated directly into the PDF processing logic.
-   **Performance and File Handling**: Files are opened and closed properly. For XLSX, `read_only=True` and `data_only=True` are used for potentially better performance with large files.
-   **Error Handling and Logging**: `try-except` blocks are used extensively to catch and log errors during file processing and OCR.
-   **Creativity in Solving the OCR Fallback**: The PDF function first attempts direct text extraction. If that fails or yields no text, it iterates through pages, converts them to images using `pdfplumber`'s `to_image()` method, and then applies `pytesseract` for OCR.

---

Thank you for considering my application!
