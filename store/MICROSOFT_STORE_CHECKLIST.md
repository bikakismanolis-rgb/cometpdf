# Microsoft Store Checklist

## Account

- Create or use a Microsoft Partner Center developer account.
- Reserve the product name `CometPDF`.

## Required URLs

- Website: `https://cometpdf.com/`
- Privacy policy: `https://cometpdf.com/privacy/`
- Support URL: `https://cometpdf.com/`

## Submission Path Options

### Option A: Win32 installer listing

This is the fastest route. Microsoft allows unpackaged Win32 apps in the Store by providing a link to the existing installer.

Use:

```text
https://cometpdf.com/downloads/CometPDFSetup-1.0.0.exe
```

This is easier to submit, but the installer still needs to be trusted and clean.

### Option B: MSIX package

This is the cleaner Store-style route. It usually provides a better install/update experience, but takes more packaging work.

Use Microsoft MSIX Packaging Tool or Visual Studio Windows Application Packaging Project.

Important: the Store/MSIX build may need to skip classic installer behaviors such as registry right-click integration.

## Assets

- App icon/logo
- At least 1 screenshot; 4 or more recommended
- Plain text description
- Privacy policy URL
- Age rating questionnaire

## Before Submission

- Verify `https://cometpdf.com` and `https://www.cometpdf.com`.
- Verify downloads work over HTTPS.
- Run the installer on a clean Windows user profile.
- Check the installer with Microsoft Defender.
- Optionally submit to VirusTotal.

