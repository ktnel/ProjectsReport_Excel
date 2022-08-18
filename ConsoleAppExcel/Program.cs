using System.Diagnostics;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
//using System.Runtime.InteropServices;
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

foreach (Process proc in finalExcelProc) {
    //If Excel process started after original collection and is still running.
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

        if (sourceSheets.Count > 0) { 
            for (int i = 1; i <= sourceSheets.Count; i++) {

                //Store current index as Worksheet.
                Excel.Worksheet srcSheet = sourceSheets[i];

                //Get Worksheet sheetName.
                string sheetName = srcSheet.Name;

                //Write data from source to destination file.
                if (sheetName.Contains("Docs")) {

                    //Active projects.
                    //Only write these colCount numbers.
                    int[] activeCols = { 1, 2, 3, 4 };

                    //Destination Worksheet.
                    Excel.Worksheet destActiveWS = destSheets[1];

                    //Write info from source to destination, only specified columns.
                    WriteData(srcSheet, destActiveWS, activeCols);

                    //Archived projects.
                    int[] archiveCols = { 1, 2, 3, 4, 5, 6, 7, 8 };
                    Excel.Worksheet destArchiveWS = destSheets[2];
                    WriteData(srcSheet, destArchiveWS, archiveCols);
                }
                //Client Hosted projects.
                else if (sheetName.Contains("Client")) {
                    int[] clientCols = { 1, 2, 3, 4, 5 };
                    Excel.Worksheet destClientWS = destSheets[3];
                    WriteData(srcSheet, destClientWS, clientCols);
                }
            }

            //Save the workbooks.
            sourceWb.Save();
            destWb.Save();

            //Close the workbooks.
            sourceWb.Close();
            destWb.Close();
        }

        //Source Workbook is empty.
        else {
            Console.WriteLine("Source Workbook is empty.");
        }
    }

    //Catch all exceptions, write to Console.
    catch (Exception ex) {
        if (ex.StackTrace != null) {
            Console.WriteLine($"{ex.Message} ({GetStackLine(ex.StackTrace)})");
        }
        else { Console.WriteLine(ex.Message); };
    }

    //Quit Excel application.
    excelApp.Quit(); 
}

///Write Excel data to destination Workbook.
static void WriteData(Excel.Worksheet srcWs, Excel.Worksheet destWs, int[] srcCols)
{
    //Get the destination Worksheet Name.
    string destSheetName = destWs.Name;

    //Variable to filter Active and Archived projects (rows). Not set on Client projects.
    string rowFilter = "";

    //If the destination sheet name contains Active or Archived, set rowFilter.
    if (destSheetName.Contains("Active")) {
        rowFilter = "Active";
    }
    else if (destSheetName.Contains("Archived")) {
        rowFilter = "Archived";
    }

    //Get the last row number.
    int lastRow = srcWs.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

    //Starting row number.
    int srcRowCount = 1;
    int destRowCount = 1;

    //Write date/time stamp to cell B1.
    destWs.Cells[1, 2] = $"Export Date: {DateTime.Now}";
    
    //Move to the second row.
    destRowCount += 1;

    //Loop until finished with last row.
    while (srcRowCount <= lastRow) {
        
        ///Use same method as "lastRow" to get cells with data. This collects 16000+ cells.
        //Get row from source worksheet.
        Excel.Range srcRow = srcWs.UsedRange.EntireRow[srcRowCount].Cells;

        //Store desired cells in list.
        List<string> frstRow = new();
        for (int i = 0; i <= srcCols.Length-1; i++) {

            //Get cell value by colCount number (index value on input srcCols array).
            frstRow.Add(srcRow[srcCols[i]].Cells.Text);
        }

        //Variable used to write a specific row or not.
        bool writeRow = true;

        //If the row is not the title row. Not used for Client projects.
        if (srcRowCount > 1 & rowFilter != "") {

            //If destination Ws Name contains rowFilter but source column 4 doesn't, skip row.
            string checkCell = frstRow[3];
            if (destSheetName.Contains(rowFilter) & checkCell != rowFilter) {
                writeRow = false;
            }
        }

        //Write data if writeRow is true.
        if (writeRow) {
            //Write row (srcRowCount) to destination file.
            for (int colCount = 1; colCount <= srcCols.Length; colCount++) {

                //Get cell value by colCount number (index value on input srcCols array).
                //-1 because srcCols array is 0-based index. Excel is 1-based index.
                Excel.Range srcCell = srcRow[srcCols[(colCount - 1)]];

                //Write the cell contents to destination cell (starting at colCount A1).
                destWs.Cells[destRowCount, colCount] = srcCell;
            }
            //Move to the next destination row.
            destRowCount++;
        }
        //Move to the next row.
        srcRowCount++;
    }
    Console.WriteLine($"{srcWs.Name} Workbook Complete");
}

static string GetStackLine(string msg)
{
    //Create a new StringBuilder
    StringBuilder strLine = new StringBuilder("Line ");

    //Isolate and collect everything after "cs:line ".
    string str = "cs:line";
    int strStart = (msg.IndexOf(str) + str.Length + 1);
    string cut = msg.Substring(strStart, (msg.Length - strStart));

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
