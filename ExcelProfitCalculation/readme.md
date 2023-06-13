---

# Project: Excel Profit Calculation

This project aims to calculate the profit based on two columns in an Excel file using VBS script.

## Step 1: Start Excel Application
```vbs
Set objExcel = CreateObject("Excel.Application")
```
In this step, we initialize the Excel application.

## Step 2: Open Excel File
```vbs
Set objWorkbook = objExcel.Workbooks.Open("C:\UpWork\ExcelProfit\test.xlsx")
```
Here, we open the Excel file that contains the data.

## Step 3: Select Worksheet
```vbs
Set objWorksheet = objWorkbook.Worksheets("Sheet1")
```
We select the worksheet where the data is located. You can change the sheet name accordingly.

## Step 4: Prepare Data Columns
```vbs
Set columnB = objWorksheet.Columns("B")
Set columnC = objWorksheet.Columns("C")
```
We define the columns B and C where the data is stored.

## Step 5: Clean Data Format
```vbs
lastRow = objWorksheet.Cells(objWorksheet.Rows.Count, columnB.Column).End(-4162).Row

For i = 1 To lastRow
    If Not IsEmpty(columnB.Cells(i)) Then
        columnB.Cells(i).Value = Replace(columnB.Cells(i).Value, ".", ",")
    End If
Next

lastRow = objWorksheet.Cells(objWorksheet.Rows.Count, columnC.Column).End(-4162).Row

For i = 1 To lastRow
    If Not IsEmpty(columnC.Cells(i)) Then
        columnC.Cells(i).Value = Replace(columnC.Cells(i).Value, ".", ",")
    End If
Next
```
In this step, we clean the data format by replacing "." with "," in columns B and C.

## Step 6: Define Data Columns and Target Column
```vbs
columnB = "B" ' Letter value of the first column
columnC = "C" ' Letter value of the second column
targetColumn = "F" ' Letter value of the target column
objWorksheet.Range(targetColumn & 1).Value = "Profit"
```
We define the column letters for column B, column C, and the target column. We also write "Profit" as the header in the target column.

## Step 7: Calculate Profit
```vbs
startRow = 2 ' Change the starting row number according to your data (2 if there is a header, 1 if there is no header)

lastRow = objWorksheet.Cells(objWorksheet.Rows.Count, columnB).End(-4162).Row ' -4162: xlUp

For i = startRow To lastRow
    cellB = columnB & i
    cellC = columnC & i
    formula = "=" & cellC & "-" & cellB
    targetCell = targetColumn & i
    objWorksheet.Range(targetCell).Formula = formula

    objWorksheet.Range(cellB).NumberFormat = "General"
    objWorksheet.Range(cellC).NumberFormat = "General"
Next
```
In this step, we calculate the profit based on the values in columns B and C. The profit formula subtracts the value in column B from the value in column C. The calculated profit is written in the target column. We also apply the "General" number format to columns B and C.

## Step 8: Save and Close Excel File
```vbs
objWorkbook.Save
objWorkbook.Close
objExcel.Quit
```
Finally, we save

 the Excel file and close the Excel application.

## Step 9: Clear Memory
```vbs
Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
```
We release the memory by clearing the variables used in the project.

---