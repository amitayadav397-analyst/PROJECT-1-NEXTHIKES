# How to add and run the VBA macros (Excel for Mac)

File to import: `BikeSharingAutomation.bas`
Workbook: `nexthikes_project_excel.xlsx`, which you will save as `.xlsm`

These steps are for Excel for Mac (2016 or newer). The same macros also work in Excel for Windows.

---

## 1. Turn on the Developer tab (one time only)

1. Open Excel.
2. In the menu bar at the top of the screen, click **Excel > Preferences...** (on newer Macs it is called **Excel > Settings...**).
3. Click **Ribbon & Toolbar**.
4. In the right-hand list (Main Tabs), tick **Developer**.
5. Click **Save** and close the window. A **Developer** tab now shows in the ribbon.

## 2. Open the VBA editor

Open the workbook, then use either of these:

- **Developer > Visual Basic**, or
- the menu bar: **Tools > Macro > Visual Basic Editor**.

## 3. Import the module

1. In the VBA editor, click **File > Import File...**
2. Choose `BikeSharingAutomation.bas` and click **Open**.
3. In the Project panel on the left, under **Modules**, you should now see **BikeSharingAutomation**. Double-click it to read the code.
4. Optional check: click **Debug > Compile VBAProject**. If nothing happens, the code has no errors.
5. Close the VBA editor (the red dot, or **Excel > Close and Return to Microsoft Excel**).

## 4. Save as a macro-enabled workbook

A normal `.xlsx` file cannot keep macros, so you must change the file type:

1. **File > Save As...**
2. For **File Format**, choose **Excel Macro-Enabled Workbook (.xlsm)**.
3. Click **Save**. From now on, use the `.xlsm` file.

## 5. Add a button on the Cover sheet

1. Go to the **Cover** sheet.
2. Click **Developer > Button** (on some versions, **Developer > Insert > Button (Form Control)**).
3. Drag a rectangle on an empty area of the sheet, for example to the right of the title.
4. The **Assign Macro** window opens. Choose **RunFullReport** and click **OK**.
5. Click the button text (or right-click > **Edit Text**) and type **Run Full Report**.
6. Click any cell to finish.

If you want more buttons, repeat the same steps for **ExportReportPDF**, **HighlightAnomalies** and **ResetFilters**.
To change a button later, hold **Control** and click it, then choose **Assign Macro...**

## 6. Run it

- Click the **Run Full Report** button, or
- go to **Developer > Macros** (or **Tools > Macro > Macros...**), pick **RunFullReport** and click **Run**.

After a few seconds a message says the report is updated. Excel then opens a new sheet, **Automated Report**. A second new sheet, **Automation Log**, records every run.

## 7. Opening the file next time

When you open the `.xlsm`, Excel asks about macros. Click **Enable Macros**, or the buttons will not work.
If you don't see the question, go to **Excel > Preferences > Security** and set macros to **Disable all macros with notification**, then open the file again.

---

## What each macro does

The data in **Final_Dataset** (table `tblFinal`) is made of formulas that link back to Dataset_A and Dataset_3. Sorting it or typing over it with a macro would break those formulas. So **every macro only reads the data** and writes its results to new sheets. The cleaning was done with formulas (see Data_Quality). The macros handle the repetitive part: recalculating, summarising, checking, reporting and exporting.

### RunFullReport (the main button)
Recalculates the whole workbook, then reads all 1,000 rows of `tblFinal` into memory (an array) and loops through them once to add up the totals for each hour and each weekday. From those totals it builds the **Automated Report** sheet, which has:
- a title and timestamp
- the key numbers: total, registered, casual, both shares, average per hour, peak hour, busiest weekday, rain impact and the anomaly count
- the hourly and weekday tables, each with a chart drawn by the macro
- the top 5 busiest hours
- data checks: row count, blanks, duplicate instants, and whether cnt = casual + registered
- seven insight sentences written from the numbers

Each run deletes the old report and builds it again, so the report always matches the current data.

### ExportReportPDF
Saves the Automated Report sheet as a PDF, for example `Automated_Report_2026-10-09_1430.pdf`, in the same folder as the workbook. If the report doesn't exist yet, it builds it first. On a Mac, Excel may ask for permission to use the folder: click **Grant Access**. If the export still fails, the macro explains why, and you can use **File > Save As > PDF** instead.

### HighlightAnomalies
For each hour of the day (0 to 23), it works out the normal demand: the mean and the standard deviation. It then gives every row a Z-score, which says how far that row is from normal. Rows where |z| is at least the threshold in **Anomalies!B4** (3 by default) are listed at the bottom of the report. Busier-than-normal rows are orange and quieter ones are blue. It does not change the Anomalies sheet, because that sheet already has its own formula-based colouring. The report's data checks confirm that the VBA count and the Anomalies sheet count agree.

### ResetFilters
Clears any filter on `tblFinal` and on the Anomalies list. It also puts the Dashboard drop-downs back to their defaults: Day type, Weather, Month and Time band to **All** (cells B6, H6, N6, T6), and the date range to **01-Jan-2011** to **14-Feb-2011** (Z6 and AF6). Use it before a demo so the Dashboard shows all the data.

### Automation Log (the WriteLog helper)
Every macro adds one line to the **Automation Log** sheet: the date and time, the macro name, and the result, which is either OK or the error message. This shows the automation ran and when.

### Custom functions: DayType, WeatherLabel, DemandLevel
These are VBA functions you can also type into a cell:
- `=DayType(H2)` gives "Weekend" or "Weekday".
- `=WeatherLabel(I2)` turns the code 1-4 into "Clear", "Cloudy/Mist", "Light Rain/Snow" or "Heavy Rain".
- `=DemandLevel(cnt, mean, sd)` gives "High", "Normal" or "Low" compared with the usual demand for that hour.

The report uses them for the top-5 list and the weekday/weekend comparison.

### VBA ideas you can point to in the code
- **Loops**: `For r = 1 To mRows`, which runs over every row.
- **Conditionals**: `If ... Then ... Else` and `Select Case`.
- **Arrays**: the table is read in one step with `DataBodyRange.Value`, which is fast.
- **A Collection**: used to find duplicate instants. It works on Mac, where `Scripting.Dictionary` does not.
- **Error handling**: `On Error GoTo Fail`. If anything goes wrong, Excel's screen updating and calculation settings are always put back, and a friendly message is shown.
- **Charts made by code**: `ChartObjects.Add`, then `SetSourceData`.

## If something goes wrong

| Problem | Fix |
|---|---|
| "Macros have been disabled" | Close the file, open it again and click **Enable Macros**. |
| Button does nothing | Check that the file is `.xlsm`, then Control-click the button > **Assign Macro** > choose RunFullReport. |
| "Column 'xxx' was not found" | A header in `tblFinal` was renamed. Change it back, or edit the name in the code. |
| PDF not saved | Save the workbook first, click **Grant Access** when Excel asks, or use **File > Save As > PDF**. |
| Report looks out of date | Click **Run Full Report** again. The sheet is rebuilt every time, so don't type on it. |
