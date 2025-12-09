# Pure Web OCR Tool 📄➡️📝

This project has been converted from a Python Flask application to a **Serverless, Client-Side Web Application**.

## 🌟 Features
- **Zero Server Required**: Runs entirely in the browser using WebAssembly.
- **Privacy Focused**: Files are processed locally on your device; nothing is uploaded to cloud.
- **Arabic Support**: Fully supports OCR for Arabic and English text.
- **Format Support**: Converts PDFs and Images to properly formatted Word (.docx) documents.

## 🚀 How to Host on GitHub Pages
1. Push this repository to GitHub.
2. Go to **Settings** > **Pages**.
3. Under **Source**, select `main` (or `master`) branch.
4. Click **Save**.
5. Your site will be live at `https://yourusername.github.io/repo-name/`.

## 🛠 Technology Stack
- **PDF.js**: For reading and rendering PDF files.
- **Tesseract.js**: For Optical Character Recognition (OCR) inside the browser.
- **docx.js**: For generating Word documents.
- **Vanilla Text Processing**: Custom logic to handle text direction and cleaning.

## 📂 Legacy Code
The original Python backend code has been moved to the `legacy_server/` directory for reference. It is not used by the live website.