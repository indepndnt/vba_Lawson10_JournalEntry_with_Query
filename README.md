# vba Infor/Lawson Excel worksheet integration Tools

This project is based on Don Peterson's standalone upload worksheets available for download on the Infor Xtreme support site.

There are currently four functions:
---

1. Single-step GL40 Journal Entry upload and GL240 report equivalent.
**`[Single_Upload_worksheet.bas, Report_worksheet.bas]`** 
2. GL Transaction Detail query with built-in AP Invoice attachment drill-around (for installations where AP invoice images are recorded with the url in an API attachment).
Also provides a link on each detail line to open the full journal entry for that line (if the Report worksheet is present).
**`[Query_worksheet.bas, (optional Report_worksheet.bas)]`**
3. GL Account Balance/Activity query for reconciliations.
**`[Balances_worksheet.bas]`**
4. Multiple GL40 Journal Entry upload for when you have several related JE's together.
Also provides a link on each JE header to open the JE report for that entry (if the Report worksheet is present).
**`[Mutli_Upload_worksheet.bas, (optional Report_worksheet.bas)]`**

Setup:
---

1. Download an appropriate version of a standalone upload workbook from the Infor site;
2. Delete all worksheets besides Sheet1 (Instructions);
3. You should have vba code in Sheet1, ThisWorkbook, and a module named modCommon_NET or similar;
4. Add selected worksheets from template workbook;
5. Insert the code from *.bas files to the referenced worksheets;
6. Assign appropriate code to ActiveX controls as listed;
7. Save.

History:
---

**02/06/2017** - Worksheet modules can now be passed between workbooks without losing functionality (provided each workbook maintains the Infor modCommon_NET module). Added multi-JE upload.

**12/19/2016** - Improved performance of query returns populating the worksheet.

**10/03/2016** - Added recognition of column types for query results (a string field with leading zeroes will retain them, a currency field will be formatted as such).

**08/18/2016** - Rewrote these functions to use the framework from the standalone upload worksheets from the Infor Xtreme Downloads section.

**05/11/2016** - Improved error handling. Added a GLTRANS query with automated AP drill-around functionality.

**04/12/2016** - Released internally.

**03/31/2016** - Major functions working including ability to add/change/delete both JE header and JE detail lines as well as produce GL240 equivalent report.

**03/27/2016** - Under development. Concept worked by driving IE from VBA to log in and interact with Lawson API.
