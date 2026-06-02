# CometPDF 1.1.0

CometPDF is a free Windows desktop app for local, no-upload PDF conversion.

Download page:
https://cometpdf.com/download/

Security and checksums:
https://cometpdf.com/security/

## Downloads

- Installer: `CometPDFSetup-1.1.0.exe`
- Portable ZIP: `CometPDF-1.1.0-portable.zip`

The installer is recommended for normal use because it supports Desktop and Start Menu shortcuts, uninstall support, and optional right-click conversion from Windows File Explorer. The Portable ZIP runs without installation.

## What's New

- Added PDF to PNG export.
- Added PDF to JPG export.
- Added PDF to TXT export for selectable PDF text.
- Added PDF to Word DOCX visual export.
- Added PDF to PowerPoint PPTX visual export.
- Added no-upload, offline-first PDF conversion pages on the public website.
- Added clearer download, security, and Windows SmartScreen information.

## Supported Conversion Workflows

CometPDF converts common files to PDF:

- Word: DOC, DOCX, RTF, ODT
- Excel: XLS, XLSX, XLSM, ODS
- PowerPoint: PPT, PPTX, ODP
- Images: JPG, JPEG, PNG, WEBP, BMP, TIFF
- Emails: MSG, EML
- Text and data: TXT, CSV, LOG, MD, HTML, XML, JSON
- Existing PDF files for copy and merge workflows

CometPDF also exports PDFs to:

- PNG
- JPG
- TXT
- DOCX
- PPTX

PDF to Word and PDF to PowerPoint in this release preserve PDF pages visually. They are useful for faithful page appearance, but they are not OCR-based fully editable reconstruction.

## Privacy

CometPDF converts files locally on Windows. It does not upload selected files to a CometPDF server during conversion.

## Windows SmartScreen Notice

CometPDF is not code-signed yet. Windows SmartScreen or Chrome may show an "unknown publisher" warning when downloading or running the installer.

This is expected for a new unsigned Windows app. The CometPDF 1.1.0 installer has 0 detections on VirusTotal across 70 antivirus engines.

To continue through SmartScreen, click:

```text
More info -> Run anyway
```

If you prefer to avoid installation, use the Portable ZIP.

## SHA256

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
