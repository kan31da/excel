
using Aspose.Cells;
using Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

const string filePath = "C:\\Users\\VladislavVladovski\\Desktop\\test.xlsx";


//Workbook wb = new Workbook(fileImport);

//// Get all worksheets
//WorksheetCollection collection = wb.Worksheets;

//// Loop through all the worksheets
//for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
//{

//    // Get worksheet using its index
//    Worksheet worksheet = collection[worksheetIndex];

//    // Print worksheet name
//    Console.WriteLine("Worksheet: " + worksheet.Name);

//    // Get number of rows and columns
//    int rows = worksheet.Cells.MaxDataRow;
//    int cols = worksheet.Cells.MaxDataColumn;

//    // Loop through rows
//    for (int i = 0; i < rows; i++)
//    {

//        // Loop through each column in selected row
//        for (int j = 0; j < cols; j++)
//        {
//            // Pring cell value
//            Console.Write(worksheet.Cells[i, j].Value + " | ");
//        }
//        // Print line break
//        Console.WriteLine(" ");
//    }
//}

Application excel = new();
Workbook wb = excel.Workbooks.Open(filePath);
Worksheet ws = wb.Worksheets[1];
object cell = ws.Cells[1, 1].value;


Console.WriteLine(cell);
