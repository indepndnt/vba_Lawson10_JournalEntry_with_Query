# vba_Lawson10_JournalEntry_with_Query
Excel VBA code for an Infor/Lawson version 10 GL40 one-step Upload plus GLTRANS query with AP attachments
---
This project is based on Don Peterson's standalone upload worksheets available for download on the Infor Xtreme support site.

There are two functions:

1. Single-step GL40 Journal Entry upload and GL240 report equivalent.
2. GL Transaction Detail query with built-in AP Invoice attachment drill-around (for installations where AP invoice images are recorded with the url in an API attachment).

# Setup:

1. Download an appropriate version of a standalone upload workbook from the Infor site;
2. Delete all worksheets besides Sheet1 (Instructions);
3. You should have vba code in Sheet1, ThisWorkbook, and a module named modCommon_NET or similar;
4. Add the worksheets from our workbook: Sheet2 (Report), Sheet3 (Upload), and Sheet4 (Query);
5. Add a module named modCommon_X and insert the code from our modCommon_X.bas;
6. Insert the code from Sheet2(Report).bas, Sheet3(Upload).bas, and Sheet4(Query).bas to the referenced worksheets;
7. Assign macros to the buttons: 'Upload': Sheet3.inUpload, 'Report': Sheet2.inJournalEditRpt, 'Query': Sheet4.inGLInvoices;
8. Save.
