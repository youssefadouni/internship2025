# DYDON AI – Internship 2025 Application

Welcome to the **DYDON AI Summer Internship 2025** repository!  
We are excited to announce that we are now accepting applications for our summer internship program. 🚀

Instead of the usual process of sending CVs and waiting, we want to **simplify and speed things up**.  
If you're truly interested in joining us, let your code speak.

---

## 📝 The Task

Your mission is to implement a **clean, efficient, and well-structured Python function** that does the following:

1. Iterates through all files in the folder `uploads/`
2. Detects and processes files of the following types:
   - `.pdf`
   - `.docx`
   - `.xlsx`
3. Extracts the full text from each file.
4. For PDFs that do not contain extractable text, apply OCR using **Tesseract OCR** via its Python wrapper (`pytesseract`).

You **do not** need to apply any AI or NLP processing to the extracted text — the goal is to prepare the content for later stages in our pipeline.

---

## ✅ Requirements

- Python 3.8+
- Use well-established libraries such as:
  - `pytesseract`
  - `pdfplumber` or `PyMuPDF`
  - `python-docx`
  - `openpyxl`
- Your solution should handle errors gracefully and log any problems.
- Focus on **clarity**, **performance**, and **extensibility**.

---

## 📬 How to Apply

- Fork this repository.
- Add your solution in a file called `extract_text_[your_full_name].py`.
- Include a brief `DEV_LOG_[your_full_name].md` in your fork explaining your approach and any dependencies.
- if you want to deploy more than one file/folder, please add your full name to the name of the file/folder
- Submit your repo by creating a **pull request** to this repository.

---

## 🎯 Evaluation Criteria

We’ll evaluate submissions based on:

- Code quality and structure
- Elegance and simplicity
- Performance and file handling
- Error handling and logging
- Creativity in solving the OCR fallback

The most **elegant and performative solution** will be selected for the internship.

---

## ❗ Final Notes

- This internship is **paid** and **hands-on** – you’ll work directly with our AI/NLP team.
- We’re looking for **curious, committed, and technically sharp minds**.
- You don’t need to send us your CV — let your code do the talking.

Good luck and have fun!

— The DYDON AI Team
