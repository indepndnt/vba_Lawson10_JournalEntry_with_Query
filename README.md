# vba_Lawson10_JournalEntry_with_Query
Excel VBA code for an Infor/Lawson version 10 GL40 one-step Upload plus GLTRANS query with AP attachments
---
This project is based on Don Peterson's standalone upload worksheets available for download on the Infor Xtreme support site.

There are three functions:

1. Single-step GL40 Journal Entry upload and GL240 report equivalent.
[Upload.bas, Report.bas] 
2. GL Transaction Detail query with built-in AP Invoice attachment drill-around (for installations where AP invoice images are recorded with the url in an API attachment).
[Query.bas]
3. GL Account Balance/Activity query for reconciliations.
[Balances.bas]

# Setup:

1. Download an appropriate version of a standalone upload workbook from the Infor site;
2. Delete all worksheets besides Sheet1 (Instructions);
3. You should have vba code in Sheet1, ThisWorkbook, and a module named modCommon_NET or similar;
4. Add selected worksheets from template workbook;
5. Insert the code from *.bas files to the referenced worksheets;
6. Assign appropriate code to ActiveX controls as listed;
7. Save.
