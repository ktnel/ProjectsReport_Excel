using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
//using System.Globalization;

//Get a list of current Excel processes.
Process[] origExcelProc = Process.GetProcessesByName("EXCEL");

//Create list of process ids.
List<int> origProcIds = new();

//If another Excel process is already running.
if (origExcelProc.Length > 0) {
    foreach (Process origProc in origExcelProc) {
        origProcIds.Add(origProc.Id);
    }
}

//Do all the Excel work.
ExcelProcess();

//Garbage Collection.
GC.Collect();
GC.WaitForPendingFinalizers();
GC.Collect();
GC.WaitForPendingFinalizers();

//Get a new list of Excel processes.
Process[] finalExcelProc = Process.GetProcessesByName("EXCEL");

//If Excel Process started after original collection and is still running.
foreach (Process proc in finalExcelProc) {
    //If new Process is running, Refresh, GC, and Kill.
    if (!origProcIds.Contains(proc.Id) && !proc.HasExited) {
        CheckFinalProcess(proc);
    }
}

///Perform all the Excel work.
static void ExcelProcess()
{
    Excel.Application excelApp = new();

    try {
        //Open source file.
        Excel.Workbooks srcBooks = excelApp.Workbooks;
        Excel.Workbook sourceWb = srcBooks.Open(@"C:\Users\kyle.nelson\Downloads\Test_BIM 360 Project List.xlsx");
        Excel.Sheets sourceSheets = sourceWb.Worksheets;

        //Open destination file.
        Excel.Workbook destWb = excelApp.Workbooks.Open(@"C:\Users\kyle.nelson\Downloads\Dest_BIM 360 Project List.xlsx");
        Excel.Sheets destSheets = destWb.Worksheets;

        //Delete contents from all destination worksheets.
        foreach (Excel.Worksheet dSheet in destSheets) {
            dSheet.UsedRange.Clear();
        }

        //If there are Worksheets in the workbook.
        if (sourceSheets.Count > 0) { 
            for (int i = 1; i <= sourceSheets.Count; i++) {

                //Store current index as Worksheet.
                Excel.Worksheet srcSheet = sourceSheets[i];

                //Get Worksheet sheetName.
                string sheetName = srcSheet.Name;

                ///Windows form to pair srcSheet with destSheet. Include filter settings.
                //Write data from source to destination file.
                if (sheetName.Contains("Docs")) {

                    //Active projects.
                    //Only write these colCount numbers.
                    int[] activeCols = { 1, 2, 3, 4, 16 };

                    //Destination Worksheet.
                    Excel.Worksheet destActiveWS = destSheets[1];

                    //Write info from source to destination, only specified columns.
                    WriteData(excelApp, srcSheet, destActiveWS, activeCols);

                    //Archived projects.
                    int[] archiveCols = { 1, 2, 3, 4, 5, 6, 7, 8, 16 };
                    Excel.Worksheet destArchiveWS = destSheets[2];
                    WriteData(excelApp, srcSheet, destArchiveWS, archiveCols);
                }
                //Client Hosted projects.
                else if (sheetName.Contains("Client")) {
                    int[] clientCols = { 1, 2, 3, 4, 5, 19 };
                    Excel.Worksheet destClientWS = destSheets[3];
                    WriteData(excelApp, srcSheet, destClientWS, clientCols);
                }
            }

            //Close the workbooks with save action.
            sourceWb.Close(false);
            destWb.Close(true);
        }

        //Source Workbook is empty.
        else {
            Console.WriteLine("Source Workbook is empty.");
        }
    }

    //Catch all exceptions, write to Console.
    catch (Exception ex) {
        if (ex.StackTrace != null) {
            //Write Exception message and line that Exception was raised on.
            Console.WriteLine($"{ex.Message} ({GetStackLine(ex.StackTrace)})");
        }
        //If no line was generated, only write Exception message.
        else { Console.WriteLine(ex.Message); };

        //Count the Workbooks that are still open.
        int wbCount = excelApp.Workbooks.Count;

        //If workbooks are still open, close without saving.
        if (wbCount > 0) {
            for (int i = 1; i <= wbCount; i++) {

                //Must be index 1 because Workbooks.Count value changes.
                //Raises exception if 'i' is greater than Workbooks.Count.
                excelApp.Workbooks[1].Close(false);
            }
            Console.WriteLine("Workbooks were closed without saving.");
        }
    }

    //Quit Excel application.
    excelApp.Quit(); 
}

///Write Excel data to destination Workbook.
static void WriteData(Excel.Application xlApp, Excel.Worksheet srcWs, Excel.Worksheet destWs, int[] srcCols)
{
    //Get the destination Worksheet Name.
    string destSheetName = destWs.Name;

    //Variable to filter Active and Archived projects (rows). Not set on Client projects.
    string? rowFilter = null;

    //If the destination sheet name contains Active or Archived, set rowFilter.
    if (destSheetName.Contains("Active")) {
        rowFilter = "Active";
    }
    else if (destSheetName.Contains("Archived")) {
        rowFilter = "Archived";
    }

    //Get the last row & column numbers that have data.
    int lastRowNum = srcWs.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
    int lastColNum = srcWs.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;

    //List to iterate the source Worksheet rows.
    List<int> iterRows = new();

    //If rowFilter was set.
    if (rowFilter != null) {

        //Create a list for rows that pass rowFilter.
        iterRows.Add(1);

        //Column used to look for rowFilter.
        int refCol = 4;

        //If refCol value matches rowFilter value, add row number to list.
        for (int i = 2; i <= lastRowNum; i++) {
            if (srcWs.Cells[i, refCol].Value == rowFilter) {
                iterRows.Add(i);
            }
        }
        //Change the last row count.
        lastRowNum = iterRows.Last();
    }

    //Source Worksheet is not being filtered.
    else {
        //List includes 1 thru lastRowNum.
        for (int i = 1; i <= lastRowNum; i++) {
            iterRows.Add(i);
        }
    }

    //First row in destination Worksheet.
    int destRowNum = 1;

    //Write date/time stamp to cell B1. Move to the second row.
    destWs.Cells[1, 2] = $"Export Date: {DateTime.Now}";
    destRowNum += 1;

    //Loop until finished with last row.
    for (int srcRowIndex = 0; srcRowIndex <= iterRows.Count - 1; srcRowIndex++) {

        //Source row number is index value from iterRows.
        int srcRowNum = iterRows[srcRowIndex];

        //Get row from source worksheet. Using first and last cell in row.
        Excel.Range firstCell = srcWs.Cells[srcRowNum, 1];
        Excel.Range lastCell = srcWs.Cells[srcRowNum, lastColNum];
        Excel.Range srcRow = srcWs.Range[firstCell, lastCell];

        //If column 1 or 3 in srcRow != n/a.
        if (srcRow[1].Text != "n/a" && srcRow[3].Text != "n/a") {
            //Write row (srcRowNum) to destination file.
            for (int colCount = 1; colCount <= srcCols.Length; colCount++) {

                //Get cell value by colCount number (index value on input srcCols array).
                //-1 because srcCols array is 0-based index. Excel is 1-based index.
                Excel.Range srcCell = srcRow[srcCols[(colCount - 1)]];

                //Write the cell contents to destination cell (starting at colCount A1).
                Excel.Range destCell = destWs.Cells[destRowNum, colCount];
                // Doesn't write cell contents using destCell is a variable.
                destWs.Cells[destRowNum, colCount] = srcCell;

                //Format destination cell using source cell format.
                destCell.HorizontalAlignment = srcCell.HorizontalAlignment;
                destCell.WrapText = srcCell.WrapText;
                destCell.Font.Bold = srcCell.Font.Bold;
                destCell.Font.Size = srcCell.Font.Size;
                destCell.Font.Color = srcCell.Font.Color;

                //Set column widths on first row after date/time stamp.
                if (destRowNum == 2) {
                    destCell.ColumnWidth = srcCell.ColumnWidth;
                }
            }
            //Move to the next destination row.
            destRowNum++;
        }
    }

    //Set destination Worksheet as the Active Worksheet.
    Excel.Worksheet activeWs = xlApp.ActiveSheet;

    //Select cell A1.
    Excel.Range cellA1 = activeWs.get_Range("A1");
    cellA1.Select();

    //Write Complete message to Console.
    Console.WriteLine($"{srcWs.Name} > {destSheetName} Workbook Complete");
}

/// Extract code line from the Exception StackTrace string.
static string GetStackLine(string msg)
{
    //Create a new StringBuilder
    StringBuilder strLine = new("Line ");

    //Isolate and collect everything after "cs:line ".
    string str = "cs:line";
    int strStart = (msg.IndexOf(str) + str.Length + 1);
    //strStart.. = strStart index to end of string.
    string cut = msg[strStart..];
    //Same as: string cut = msg.Substring(strStart, (msg.Length - strStart));

    //Convert each char into int and add to StringBuilder.
    //  once a number isn't found, break the foreach loop.
    foreach (char c in cut) {
        bool success = int.TryParse(c.ToString(), out int number);
        if (success) {
            strLine.Append(c);
        }
        else { break; }
    }

    return strLine.ToString();
}

/// If Excel Processes started in this application are still running, stop them.
static void CheckFinalProcess(Process process)
{
    int counter = 2;

    //Refresh # times, then kill it.
    for (int i = 0; i <= counter; i++) {
        //If Process.HasExited is false.
        if (i < counter) {
            //Discard cached info about process.
            process.Refresh();

            //Garbage Collection.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //Write message to Console.
            Console.WriteLine("Excel Process is still running.");
            Thread.Sleep(1000);
        }

        //On the last time, kill the process.
        else if (i == counter) {
            //Kill the process.
            process.Kill();
            Console.WriteLine("Excel Process has been terminated.");
        }
    }
}

///Create a new Worksheet.
static Worksheet CreateWorksheet(Workbook hostWb, string wbName)
{
    Worksheet newWorksheet = hostWb.Sheets.Add();
    newWorksheet.Name = wbName;

    return newWorksheet;
}
