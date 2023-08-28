//using System.Globalization;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


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
foreach (Process proc in finalExcelProc)
{
    //If new Process is running, Refresh, GC, and Kill.
    if (!origProcIds.Contains(proc.Id) && !proc.HasExited) {
        CheckFinalProcess(proc);
    }
}

///Perform all the Excel work.
static void ExcelProcess()
{
    //Application excelApp = new();
    string sourceFilePath = @"C:\Users\kyle.nelson\OneDrive - THERMA CORPORATION\Documents\Reports\BI_BIM 360 Project List.xlsx";
    string destinationFilePath = @"C:\Users\kyle.nelson\THERMA CORPORATION\Obernel BIM - Reports\BIM 360 Project Status Report.xlsx";

    using (SpreadsheetDocument sourceDocument = SpreadsheetDocument.Open(sourceFilePath, false))
    using (SpreadsheetDocument destDocument = SpreadsheetDocument.Open(destinationFilePath, true))
    {
        try
        {
            // Source file
            WorkbookPart srcWbPart = sourceDocument.WorkbookPart;
            Workbook srcWb = srcWbPart.Workbook;
            Sheets sourceSheets = srcWb.Sheets;

            if (sourceSheets.Count() == 0)
            {
                throw new ApplicationException("Source Spreadsheet is empty");
            }

            //Destination file
            WorkbookPart destWbPart = destDocument.WorkbookPart;
            Workbook destWb = destWbPart.Workbook;
            Sheets destSheets = destWb.Sheets;


            //Delete contents from all destination worksheets.
            foreach (Sheet dSheet in destSheets)
            {
                WorksheetPart destWsPart = (WorksheetPart)destDocument.WorkbookPart.GetPartById(dSheet.Id.Value);
                SheetData destSheetData = destWsPart.Worksheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = destWsPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

                foreach (Row row in rows)
                {
                    row.Remove();
                }
            }

            foreach (Sheet srcSheet in sourceSheets)
            {
                WorksheetPart srcWsPart = (WorksheetPart)sourceDocument.WorkbookPart.GetPartById(srcSheet.Id.Value);
                SheetData srcSheetData = srcWsPart.Worksheet.GetFirstChild<SheetData>();
                string srcSheetName = srcSheet.Name;

                if (srcSheetName.Contains("Docs"))
                {
                    int[] activCols = { 1, 2, 3, 4, 16, 17 };
                    List<int> activeRowIndices = new List<int>();
                    List<int> archiveRowIndices = new List<int>();
                    IEnumerable<Row> srcRows = srcSheetData.Elements<Row>();
                    
                    foreach (Row row in srcRows)
                    {
                        try
                        {
                            Cell cell = row.Elements<Cell>().ElementAt(3);
                            string cellValue = GetCellValue(cell, srcWbPart);

                            if (cellValue == "Active")
                            {
                                activeRowIndices.Add(srcRows.ToList().IndexOf(row));
                            }
                            else if (cellValue == "Archived")
                            {
                                archiveRowIndices.Add(srcRows.ToList().IndexOf(row));
                            }
                        }
                        catch (Exception ex)
                        {
                            string msg = ex.Message;
                            int rowNum = srcRows.ToList().IndexOf(row);

                            if (rowNum > srcRows.Count() +1)
                            {
                                throw;
                            }
                        }
                    }

                    // Destination Worksheet for Active projects.
                    Sheet destActiveWs = destWbPart.Workbook.Descendants<Sheet>().ElementAt(0);
                    WriteDocsData(srcWsPart, srcSheetData, destActiveWs, activeRowIndices, activCols);

                    // Destination Worksheet for Archived projects.
                    Sheet destArchiveWs = destWbPart.Workbook.Descendants<Sheet>().ElementAt(1);
                    WriteData(srcSheetData, destArchiveWs, activCols);
                }

                else if (srcSheetName.Contains("Client"))
                {
                    int[] clientCols = { 1, 2, 3, 4, 5, 6, 19, 20 };

                    // Destination Worksheet for Client Hosted projects.
                    Sheet destClientWs = destWbPart.Workbook.Descendants<Sheet>().ElementAt(2);
                    WriteData(srcSheetData, destClientWs, clientCols);
                }
            }
            //Close the workbooks with save action.
            sourceDocument.Close();
            destDocument.Close();

        }
        //Catch all exceptions, write to Console.
        catch (Exception ex)
        {
            if (ex.StackTrace != null)
            {
                //Write Exception message and line that Exception was raised on.
                Console.WriteLine($"{ex.Message} ({GetStackLine(ex.StackTrace)})");
            }
            //If no line was generated, only write Exception message.
            else
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

static void WriteDocsData(WorksheetPart wsPart, SheetData sheetData, Sheet destSheet, List<int> rowIndices, int[] colNumbers)
{
    SheetData destSheetData = destWsPart.Worksheet.GetFirstChild<SheetData>();
    IEnumerable<Row> rows = destWsPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();
}

///Write Excel data to destination Workbook.
static void WriteData(SheetData srcSheetData, Sheet destSheet, int[] srcDataColsNums)
{
    //Get the destination Worksheet Name.
    string destSheetName = destSheet.Name;

    //Variable to filter Active and Archived projects (rows). Not set on Client projects.
    string rowFilter;

    //If the destination sheet name contains Active or Archived, set rowFilter.
    if (destSheetName.Contains("Active")) {
        rowFilter = "Active";
    }
    else if (destSheetName.Contains("Archived")) {
        rowFilter = "Archived";
    }

    //Get the last row & column numbers that have data.
    IEnumerable<Row> srcRows = srcSheetData.Elements<Row>();
    int lastRowNum = srcRows.Count();

    IEnumerable<Column> srcCols = srcSheetData.Elements<Column>();
    int lastColNum = srcCols.Count();

    /*
    //List to iterate the source Worksheet rows.
    List<int> iterRows = new();
    
    //If rowFilter was set.
    if (rowFilter != null)
    {
        //Create a list for rows that pass rowFilter.
        iterRows.Add(1);

        //Column used to look for rowFilter.
        int refCol = 4;

        //If refCol value matches rowFilter value, add row number to list.
        for (int i = 2; i <= lastRowNum; i++)
        {
            if (srcSheetData.Cells[i, refCol].ToString() == rowFilter)
            {
                iterRows.Add(i);
            }
        }
        //Change the last row count.
        lastRowNum = iterRows.Last();
    }

    //Source Worksheet is not being filtered.
    else
    {
        //List includes 1 thru lastRowNum.
        for (int i = 1; i <= lastRowNum; i++)
        {
            iterRows.Add(i);
        }
    }

    //First row in destination Worksheet.
    int destRowNum = 1;
    
    //Write date/time stamp to cell B1. Make text bold/red.
    Range timeStampCell = (Range)destSheet.Cells[1, 2];
    timeStampCell.Value = $"Export Date: {DateTime.Now}"; //destWs.Cells[1, 2] = $"Export Date: {DateTime.Now}";
    timeStampCell.Font.Bold = true;
    timeStampCell.Font.Color = System.Drawing.Color.Red;

    //Move iterator to the next(2nd) row.
    destRowNum += 1;

    //Loop until finished with last row.
    for (int srcRowIndex = 0; srcRowIndex <= iterRows.Count - 1; srcRowIndex++)
    {
        //Source row number is index value from iterRows.
        int srcRowNum = iterRows[srcRowIndex];

        //Get row from source worksheet. Using first and last cell in row.
        Range firstCell = (Range)srcSheet.Cells[srcRowNum, 1];
        Range lastCell = (Range)srcSheet.Cells[srcRowNum, lastColNum];
        Range srcRow = srcWs.Range[firstCell, lastCell];

        //If column 1 or 3 in srcRow != n/a.
        if (srcSheet[1].ToString() != "n/a" && srcSheet[3].ToString() != "n/a")
        {
            //Write row (srcRowNum) to destination file.
            for (int colCount = 1; colCount <= srcCols.Length; colCount++)
            {
                //Get cell value by colCount number (index value on input srcCols array).
                //-1 because srcCols array is 0-based index. Excel is 1-based index.
                Range srcCell = (Range)srcSheet[srcCols[(colCount - 1)]];

                //Write the cell contents to destination cell (starting at colCount A1).
                Range destCell = (Range)destSheet.Cells[destRowNum, colCount];

                // Won't write cell contents using destCell as a variable.
                destSheet.Cells[destRowNum, colCount] = srcCell;

                //Format destination cell using source cell format.
                destCell.HorizontalAlignment = srcCell.HorizontalAlignment;
                destCell.WrapText = srcCell.WrapText;
                destCell.Font.Bold = srcCell.Font.Bold;
                destCell.Font.Size = srcCell.Font.Size;
                //destCell.Font.Color = srcCell.Font.Color;

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
    //Worksheet activeWs = xlApp.ActiveSheet;
    Worksheet activeWs = (Worksheet)xlApp.ActiveSheet;

    //Select cell A1.
    Range cellA1 = activeWs.get_Range("A1");
    cellA1.Select();
    */
    //Write Complete message to Console.
    Console.WriteLine($"{srcSheetData.Ancestors<Sheet>().FirstOrDefault().Name} > {destSheetName} Workbook Complete");////////////////////////////////////////////////////
}

static string GetCellValue(Cell cell, WorkbookPart workbookPart)
{
    string value = cell.InnerText;

    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
    {
        var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

        if (stringTable != null)
        {
            value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
        }
    }

    return value;
}


/// Extract code line from the Exception StackTrace string.
static string GetStackLine(string msg)
{
    //Create a new StringBuilder
    StringBuilder strLine = new("Line ");

    //Isolate and collect everything after "cs:line ".
    string str = "cs:line";
    int strStart = (msg.IndexOf(str) + str.Length + 1);
    
    string cut = msg[strStart..];
    //strStart.. = strStart index to end of string.
    //Same as: string cut = msg.Substring(strStart, (msg.Length - strStart));

    //Convert each char into int and add to StringBuilder.
    //  once a number isn't found, break the foreach loop.
    foreach (char c in cut)
    {
        bool success = int.TryParse(c.ToString(), out int number);
        if (success){
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
    for (int i = 0; i <= counter; i++)
    {
        //If Process.HasExited is false.
        if (i < counter)
        {
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
        else if (i == counter)
        {
            //Kill the process.
            process.Kill();
            Console.WriteLine("Excel Process has been terminated.");
        }
    }
}

/*///Create a new Worksheet.
static Worksheet CreateWorksheet(Workbook hostWb, string wbName)
{
    Worksheet newWorksheet = hostWb.Sheets.Add();
    newWorksheet.Name = wbName;

    return newWorksheet;
}*/
