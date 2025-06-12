# Dedupe Application

1. Extract Origination Report [Report Date].
2. Open the extracted Origination Report File.
3. Set filter.
4. Filter Column D [New].
5. Filter Column H unselect [Voided].
6. Copy All Data.
7. Open [Dups] file and select [Raw New and Approved only] worksheet.
8. Paste copied data in the last empty row. Paste all.
9. Apply formula in column V and W.
10. Go to worksheet [Unique Counter].
11. Delete all previous data if exists.
12. Go back to worksheets [Raw New and Approved only] filter the current report date. Copy all visible data.
13. Go to worksheet [Unique Counter] paste values and formatting.
14. Apply condtional formatting in Column I [Email Address] highlight duplicate values.
15. Check for duplicate values, use filter by color.
16. [If there is a duplicate go to this link.](#manage-duplicates)
17. Go to worksheet [Summary] fill columns B and D. If no duplicates found place 0 value.
18. Go to worksheets [Unique Counter] count all applications. Place the count in Summary [Total Approved New Customers].
19. Go back to worksheets [**Raw New and Approved only**] copy data and paste it in the last empty row of worksheets [**Dec-JanOrginated New and ap**].
20. Fix the pivot [**Data Source**] for both worksheets [**Raw New and Approved only**]  &  [**Dec-JanOrginated New and ap**].
21. Refresh All
22. [**Raw New and Approved only**] change Pivot filter on [**Date**]
23. [**Raw New and Approved only**] populate formula in Pivot table.
24. [**Dec-JanOrginated New and ap**] Date filter should be selected [**All**] and current report date is checked.
25. Workhsheets [**Raw New and Approved only**] filter 2.
26. Count the application with 2, and fill worksheets [**Summary**] column C [**Total Count of Dups May June Running**].
27. Transfer Summary data ito [**Summary of Dups**] file and send to Ms. Eve.



## Manage Duplicates
1. If there is a duplicate in [**Email At**] column I, filter by color [**Red**] to select duplicates only.
2. Copy the data with duplicates and paste it in worksheets [**Sheet3**].
3. Go to worksheets [**Pivot**] fix the data source [**Change Data Source**]
4. If [Pivot] has data, fill the [**Summary**] with the count of duplicates.
5. Note exclude the unique in duplicates in [**Count of Customers with multiple loans (Unique)**] column B. In [**Total Approved with multiple loans**] column D, input what is reflecting in the Pivot summary.
5. For example 1 application has duplicates, [**Count of Customers with multiple loans (Unique)**] value is 1, [**Total Approved with multiple loans**] value is 2.
