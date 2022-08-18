using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

//Get a list of current Excel processes.
Process[] origExcelProc = Process.GetProcessesByName("EXCEL");

//Create list of process ids.
List<int> origProcIds = new();

//If another Excel process is already running.
if (origExcelProc.Length > 0) {

    //Add process Id to the list.
    for (int i = 0; i <= origExcelProc.Length -1; i++) {
        origProcIds.Add(origExcelProc[i].Id);
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

for (int i = 0; i <= finalExcelProc.Length -1; i++) {
    //Store the iteration in a variable.
    Process proc = finalExcelProc[i];

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

        foreach (Excel.Worksheet dSheet in destSheets) {
            dSheet.UsedRange.Clear();
        }

        for (int i = 1; i <= sourceSheets.Count; i++) {

            //Store current index as Worksheet.
            Excel.Worksheet sheet = sourceSheets[i];

            //Get Worksheet sheetName.
            string sheetName = sheet.Name;

            //Write date from source to destination file.
            if (sheetName.Contains("Docs")) {
                int[] activeCols = { 1, 2, 3, 4 };
                WriteHeader(sheet, destSheets[1], activeCols);

                int[] archiveCols = { 1, 2, 3, 4, 5, 6, 7, 8 };
                WriteHeader(sheet, destSheets[2], archiveCols);
            }
            else if (sheetName.Contains("Client")) {
                int[] clientCols = { 1, 2, 3, 4, 5 };
                WriteHeader(sheet, destSheets[3], clientCols);
            }
        }

        //Save the workbooks.
        sourceWb.Save();
        destWb.Save();

        //Close the workbooks.
        sourceWb.Close();
        destWb.Close();
    }
    catch (Exception ex) {
        Console.WriteLine(ex.Message);
    }

    //Quit Excel application.
    excelApp.Quit(); 
}

///Write Excel data to destination Workbook.
static void WriteHeader(Excel.Worksheet srcWs, Excel.Worksheet destWs, int[] cols)
{
    //Write date/time stamp to cell B1.
    destWs.Cells[1, 2] = $"Export Date: {DateTime.Now}";
    
    //Get the last row number.
    int lastRow = srcWs.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

    //Starting row number. 2 because date/time is written to row 1.
    int rowCount = 2;
    
    //Loop until finished with last row.
    while (rowCount <= lastRow) {

        //Get row from source worksheet. Current row is one less than starting row.
        Excel.Range srcRow = srcWs.UsedRange.EntireRow[rowCount - 1].Cells;
        
        //Write row (rowCount) to destination file.
        for (int column = 1; column <= cols.Length; column++) {

            //Get cell value by column number (index value on input cols array).
            Excel.Range srcCell = srcRow[cols[(column - 1)]];

            //Write the cell contents to destination cell (starting at column A1).
            destWs.Cells[rowCount, column] = srcCell;
        }
        
        //Move to the next row.
        rowCount++;
    }
    Console.WriteLine("Write Done");
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
