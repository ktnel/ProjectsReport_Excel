using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

//Get a list of current Excel process Ids.
Process[] origExcelProc = Process.GetProcessesByName("EXCEL");
List<int> origProcId = new List<int>();
foreach (Process p in origExcelProc) {
    origProcId.Add(p.Id);
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
    if (!origProcId.Contains(proc.Id) && !proc.HasExited) {
        CheckFinalProcess(proc);
    }
}


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
            DocProjects(sheet, destSheets[1]);
        }
        else if (sheetName.Contains("Client")) {
            ClientProjects(sheet, destSheets[2]);
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


static void DocProjects(Excel.Worksheet srcWs, Excel.Worksheet destWs)
{
    Excel.Range firstRow = srcWs.UsedRange.EntireRow[1].Cells;
    int lastCol = srcWs.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
    int lastRow = srcWs.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

    for (int i = 1; i <= lastCol; i++) {
        //Write first row from source to destination worksheet.
        destWs.Cells[1, i] = firstRow[i];
    }
}

static void ClientProjects(Excel.Worksheet srcWs, Excel.Worksheet destWs)
{

}

///<summary>
/// If Excel Processes started in this application are still running, stop them.
/// </summary>
/// <param name="origIds"></param>
static void CheckFinalProcess(Process process)
{
    //Refresh 5 times, then kill it.
    for (int i = 0; i <= 5; i++) {
        //If Process.HasExited is false.
        if (i < 5) {
            //Discard cached info about process.
            process.Refresh();
            //Write message to Console.
            Console.WriteLine("Excel Process is still running.");
            Thread.Sleep(1000);
        }

        //On the 6th time, kill the process.
        else if (i == 5) {
            //Kill the process.
            process.Kill();
            Console.WriteLine("Excel Process has been terminated.");
        }
    }
}
