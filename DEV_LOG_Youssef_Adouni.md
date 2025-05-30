Written on 30/05/2025

I built a Python script that extracts text from PDFs, Word docs (DOCX), and Excel files (XLSX) into clean text files.
For PDFs, it grabs regular text first, then uses pytesseract with Pillow to read text from any images.
For Word files, it extracts both paragraphs and tables, and for Excel, it saves each sheet's data with tabs separating columns. 
Although pdfplumber and PyMuPDF were the suggested libraries to use in repository description, I have found much better results using pypdf. That is why the script relies on pypdf for PDFs, python-docx for Word, and openpyxl for Excel, saving everything in an organized "Youssef_Output" folder. 
I aimed to keep the code simple yet as effective as possible - it processes each file type separately while maintaining as much as possible the original structure, making it easy to expand later for other formats or improvements. The whole thing runs automatically on any files placed in the "uploads" folder.