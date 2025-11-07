I am using Excel 2016 (Polish version) — so I do NOT have functions like XLOOKUP, FILTER, MAXIFS or TEXTJOIN.
I have a table with sales data on sheet "Clean Data", in the range C5:G35.
Row 5 contains headers:
C: #
D: Salesperson
E: Region
F: Product
G: Sales Amount
There are 3 regions (but may change in the future).
✅ I need a formula that returns:
the best salesperson for each region
if there is a tie (same highest Sales Amount), return all tied names
The region to check is in cell O9.
The result should appear in P9.
Please:
give me a working formula for Excel 2016 PL language
if array formula is needed, tell me to confirm with Ctrl+Shift+Enter
If different options exist (helper columns, VBA, PowerQuery), mention them, but focus on formula-first
Provide:
✅ Formula for top salesperson
✅ Formula for tie case (if possible)
✅ Alternative solutions if Excel 2016 formula can’t list all ties
