# CometPDF

CometPDF is a small Windows desktop app for converting everyday files to PDF locally.

It is built for users who want a simple, private way to convert Office documents, images, emails, text files, existing PDFs, and PDF exports without uploading files to an online service.

## Quick Start

1. Run `CometPDFSetup-1.1.0.exe`, or open the portable `CometPDF.exe`.
2. Add files with the file picker or drag them into the app.
3. Choose an output folder if needed.
4. Click `Convert to PDF`.
5. The PDF is created locally on the computer.

Settings are saved in `cometpdf_settings.json`.

## PDF Export Tools

CometPDF can also export PDF files to other formats:

- PDF to PNG images
- PDF to JPG images
- PDF to TXT
- PDF to PowerPoint, with each PDF page as a slide image
- PDF to Word, with each PDF page preserved visually inside a DOCX file

PDF to Word in this release prioritizes visual fidelity over editability. Scanned PDFs may need OCR, which is not included yet.

## Windows Right-Click Conversion

The installer can add `Convert to PDF` to the Windows right-click menu.

When using right-click conversion, CometPDF saves the PDF next to the original file and opens it automatically.

## Release Files

- App: `dist\CometPDF.exe`
- Installer: `release\CometPDFSetup-1.1.0.exe`
- Portable ZIP: `release\CometPDF-1.1.0-portable.zip`
- Website ZIP: `release\CometPDF-website.zip`

## Supported Files

- Images: `.jpg`, `.jpeg`, `.png`, `.bmp`, `.tif`, `.tiff`, `.webp`
- HEIC/HEIF: `.heic`, `.heif` when image decoder support is available
- Text/data: `.txt`, `.csv`, `.log`, `.md`, `.html`, `.htm`, `.xml`, `.json`
- Email: `.eml`, `.msg`
- Existing PDFs: `.pdf`
- Office files: `.doc`, `.docx`, `.rtf`, `.odt`, `.xls`, `.xlsx`, `.xlsm`, `.ods`, `.ppt`, `.pptx`, `.odp`

## Office Conversion

For Word, Excel, PowerPoint, Outlook MSG files, and LibreOffice/OpenOffice formats, CometPDF first tries native export through Microsoft Office or LibreOffice when available.

That gives the closest result to printing or exporting from the original application.

If native export is not available, CometPDF uses built-in fallback renderers for supported file types where possible.

## Website

The static website is in `website/`.

Production domain:

```text
https://cometpdf.com
```

For Vercel, set the project root directory to:

```text
website
```

## Command Line

```powershell
.\dist\CometPDF.exe --convert "example.txt" --output "example.pdf"
.\dist\CometPDF.exe --convert "example.pdf" --mode pdf_png --output "example.png"
.\dist\CometPDF.exe --convert "example.pdf" --mode pdf_txt --output "example.txt"
.\dist\CometPDF.exe --convert "example.pdf" --mode pdf_pptx --output "example.pptx"
.\dist\CometPDF.exe --convert "example.pdf" --mode pdf_docx --output "example.docx"
```
