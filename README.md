# CometPDF

CometPDF is a free Windows desktop app for local, no-upload PDF conversion and PDF export.

It converts everyday files to PDF and also exports PDFs back to useful formats such as images, text, Word, and PowerPoint. CometPDF is built for people who want a simple PDF tool for private files, office work, email archiving, batch conversion, and quick document delivery without sending files to an online converter.

## Quick Links

- Website: https://cometpdf.com
- Download: https://cometpdf.com/download/
- Security and checksums: https://cometpdf.com/security/
- Privacy policy: https://cometpdf.com/privacy/
- Offline PDF converter page: https://cometpdf.com/offline-pdf-converter/
- No-upload PDF converter page: https://cometpdf.com/no-upload-pdf-converter/
- Outlook MSG and EML to PDF: https://cometpdf.com/msg-to-pdf/

## Why CometPDF?

- No uploads: files are converted on your Windows computer.
- Offline-first: normal conversion works without an internet connection.
- Batch-friendly: add many files and convert them in one run.
- Optional merge: combine PDF outputs into one PDF.
- Installer or portable: choose setup integration or a no-install ZIP.
- Windows Explorer workflow: installer can add right-click Convert to PDF.
- Security page: SHA256 checksums and VirusTotal scan notes are published on the website.

## Download

Current public release: CometPDF 1.1.0

- Installer: https://cometpdf.com/downloads/CometPDFSetup-1.1.0.exe
- Portable ZIP: https://cometpdf.com/downloads/CometPDF-1.1.0-portable.zip

The installer is recommended for normal use because it supports Desktop and Start Menu shortcuts, uninstall support, and optional right-click conversion. The Portable ZIP is useful when you want to run CometPDF without installation.

## Windows SmartScreen Notice

CometPDF is not code-signed yet, so Windows SmartScreen or Chrome may show an "unknown publisher" warning when downloading or running the installer. This is expected for a new unsigned Windows app.

To proceed through SmartScreen, click `More info` and then `Run anyway`. If you prefer to avoid installation, use the Portable ZIP.

The current public installer has 0 detections on VirusTotal across 70 antivirus engines:

https://www.virustotal.com/gui/file/947e4f4bcd01c02f95d9448f774b0f8487d9c178ad1e39e1c750ee7f25e6dbcf/detection

## SHA256 Checksums

Installer:

```text
947E4F4BCD01C02F95D9448F774B0F8487D9C178AD1E39E1C750EE7F25E6DBCF
```

Portable ZIP:

```text
ABB4B16F669FE1CDE37F6B5BDBBA04D63660C1FEB30BF3BBEE50493DC89367ED
```

Standalone EXE:

```text
27C3D6040607B2D3B10C87E161304BECCAF0AFA24463D469AC2C55348BF504F5
```

## Supported Files

CometPDF converts these file types to PDF:

- Word documents: `.doc`, `.docx`, `.rtf`, `.odt`
- Excel spreadsheets: `.xls`, `.xlsx`, `.xlsm`, `.ods`
- PowerPoint presentations: `.ppt`, `.pptx`, `.odp`
- Images: `.jpg`, `.jpeg`, `.png`, `.bmp`, `.tif`, `.tiff`, `.webp`
- Email files: `.eml`, `.msg`
- Text and data files: `.txt`, `.csv`, `.log`, `.md`, `.html`, `.htm`, `.xml`, `.json`
- Existing PDFs: `.pdf`

HEIC and HEIF support depends on image decoder support being available on the machine.

## PDF Export Tools

CometPDF 1.1.0 can also export PDFs to:

- PNG images
- JPG images
- TXT text files
- Word DOCX, with each PDF page preserved visually
- PowerPoint PPTX, with each PDF page as a slide image

PDF to Word and PDF to PowerPoint in this release prioritize visual fidelity over full editability. OCR and advanced PDF page tools are planned for a future Pro release.

## Popular Workflows

- Convert Word, Excel, and PowerPoint files to PDF without uploading.
- [Convert Outlook MSG and EML email files to PDF offline](https://cometpdf.com/msg-to-pdf/) for archiving.
- Convert images and scans to PDF, including multi-page TIFF files.
- Batch convert many files and optionally merge the resulting PDFs.
- Export PDF pages to PNG or JPG images.
- Extract selectable PDF text to TXT.
- Create visual DOCX or PPTX files from PDF pages.

## Office Conversion

For Word, Excel, PowerPoint, Outlook MSG files, and LibreOffice/OpenOffice formats, CometPDF first tries to use native export through Microsoft Office, Outlook, or LibreOffice when available. This gives the closest result to printing or exporting from the original application.

If native export is not available, CometPDF uses built-in fallback renderers for supported file types where possible.

## Command Line

```powershell
.\dist\CometPDF.exe --convert "example.txt" --output "example.pdf"
.\dist\CometPDF.exe --convert "example.pdf" --mode pdf_png --output "example.png"
.\dist\CometPDF.exe --convert "example.pdf" --mode pdf_txt --output "example.txt"
.\dist\CometPDF.exe --convert "example.pdf" --mode pdf_pptx --output "example.pptx"
.\dist\CometPDF.exe --convert "example.pdf" --mode pdf_docx --output "example.docx"
```

## Website

The static website lives in `website/`.

For Vercel, set the project root directory to:

```text
website
```

Production domain:

```text
https://cometpdf.com
```

## Privacy

CometPDF is designed to convert files locally on your Windows computer. Version 1.1.0 does not include analytics, advertising SDKs, tracking pixels, accounts, cloud sync, or telemetry.

Privacy policy: https://cometpdf.com/privacy/

## Roadmap

Planned future features include OCR, PDF compression, watch folders, stronger PDF page tools, and signed distribution through trusted Windows channels.
