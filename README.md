# powerJobs-CreatePDF-Watermark
Sample jobs to Create PDFs and add Watermarks with settings. Vault 2026 tested

*Windows PowerShell • coolOrange powerJobs Client • coolOrange powerJobs Processor*

Create production-ready PDFs directly from your Vault drawings, with automatic watermarking for Work-in-Progress or Review states (non-released states) to prevent premature use.  
Finished PDFs are routed to Vault and/or a network share, providing a reliable, lifecycle-aware publishing pipeline that improves visibility while safeguarding data integrity.

---

## Disclaimer

THE SAMPLE CODE IN THIS REPOSITORY IS PROVIDED “AS IS” WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  
USE AT YOUR OWN RISK. NO SUPPORT IS PROVIDED.

---

## Description

This sample implements an automated PDF publishing workflow for Autodesk Vault Professional using coolOrange powerJobs. When triggered on IDW or DWG drawings, it:

- Exports a PDF using Inventor/AutoCAD via powerJobs.  
- Watermarks the PDF with the file’s lifecycle **State** (e.g., *WIP*, *In Review*) when the drawing is **not** a released revision.  
- Adds the PDF back to Vault (as a *DesignVisualization* file), optionally attaches it to the source file, and/or copies it to a network folder.  
- Supports fast-open for released drawings to speed up processing.  

The job is intended to run from a lifecycle-state change or similar trigger for **FILE** entities.

<img width="982" height="991" alt="image" src="https://github.com/user-attachments/assets/4ffbf432-8bf5-4afb-94b3-59994b7bd1cf" />

---

## Prerequisites

### Versions

- Autodesk Vault Professional (Client + Job Processor)  
- coolOrange powerJobs Processor **v26.0.6 or later** (required for the built-in `Add-PDFWatermark` cmdlet)  
- coolOrange powerJobs Client **v26.x** (or compatible with your Processor)  
- Autodesk Inventor and/or AutoCAD installed on the Job Processor machine (for publishing)  

### Environment

The Job Processor must have access to:

- Vault Server  
- The configured Vault folder for output (if used)  
- The configured network share (if used)  

---

## Files & Structure

Place the script in the standard powerJobs customizations path:

```text
C:\ProgramData\coolOrange\Client Customizations\Jobs\CreatePdfWithWatermark.ps1
