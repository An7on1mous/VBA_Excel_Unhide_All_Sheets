# VBA_Excel_Unhide_All_Sheets
Rather simple VBA code to unhide all worksheets in an Excel workbook. 

Sometimes we find files with a large number of hidden worksheets and unhiding them all can only be done one by one. 
Instructions
You can achieve this by following these steps:

Step 1: Open Visual Basic by pressing alt + F8 or going to Developer, Visual Basic.

Step 2: Add a new module if there is not one already. 

Step 3: Enter or copy and paste the following code:

    Sub UnhideAllSheets()

        Dim ws As Worksheet

        For Each ws In ActiveWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next ws

    End Sub

Step 4: Close Visual Basic

Step 5: Select the workbook to be processed.

Step 6: Press F5 to open the ‘run macro’ prompt  as shown in the image to the right.

Step 7: Select UnhideAllSheets and click Run.

Step 8: To save the macro file, save it as .xlsm or better, .xlsb.
