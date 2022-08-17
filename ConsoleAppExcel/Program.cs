using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

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
    
    //Open source file.
    Excel.Workbooks srcBooks = excelApp.Workbooks;
    Excel.Workbook sourceWb = srcBooks.Open(@"C:\Users\kyle.nelson\Downloads\Test_BIM 360 Project List.xlsx");
    Excel.Sheets sourceSheets = sourceWb.Worksheets;

    //Open destination file.
    Excel.Workbook destWb = excelApp.Workbooks.Open(@"C:\Users\kyle.nelson\Downloads\Dest_BIM 360 Project List.xlsx");
    Excel.Sheets destSheets = destWb.Worksheets;

    for (int i = 1; i <= sourceSheets.Count; i++) {

        //Store current index as Worksheet.
        Excel.Worksheet sheet = sourceSheets[i];

        //Get Worksheet sheetName.
        string sheetName = sheet.Name;

        //Write date from source to destination file.
        if (sheetName.Contains("Docs")) {
            WriteDocsHeader(sheet, destSheets[1]);
        }
        else if (sheetName.Contains("Client")) {
            WriteClientHeader(sheet, destSheets[2]);
        }
    }

    //Save the workbooks.
    sourceWb.Save();
    destWb.Save();

    //Close the workbooks.
    sourceWb.Close();
    destWb.Close();
    
    //Quit Excel application.
    excelApp.Quit();
    
}

///Write specific Docs header columns to destination file.
static void WriteDocsHeader(Excel.Worksheet srcWs, Excel.Worksheet destWs)
{
    //Get the last column that has data.
    int lastCol = srcWs.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;

    //Get the first row.
    Excel.Range firstRow = srcWs.UsedRange.EntireRow[1].Cells;

    //Write first row from source to destination worksheet.
    for (int i = 1; i <= lastCol; i++) {
        destWs.Cells[1, i] = firstRow[i];
    }
}

///Write specific Client header columns to destination file.
static void WriteClientHeader(Excel.Worksheet srcWs, Excel.Worksheet destWs)
{

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
