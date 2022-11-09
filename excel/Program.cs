//using Microsoft.Office.Interop.Excel;
//using Workbook = Microsoft.Office.Interop.Excel.Workbook;
//using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

using IronXL;

WorkBook wb = WorkBook.Create(ExcelFileFormat.XLSX);

wb.SaveAs("Budget.xlsx");

WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
var sheet = workbook.CreateWorkSheet("Result");

// Set Cell Values Manually
sheet["A1"].Value = "Object Oriented Programming";
sheet["B1"].Value = "Data Structure";
sheet["C1"].Value = "Database Management System";
sheet["D1"].Value = "Agile Development";
sheet["E1"].Value = "Software Design and Architecture";
sheet["F1"].Value = "Software Requirement Engineering";
sheet["G1"].Value = "Computer Programming";
sheet["H1"].Value = "Software Project Management";
sheet["I1"].Value = "Software Construction";
sheet["J1"].Value = "Software Quality Engineering";
sheet["K1"].Value = "Software ReEngineering";
sheet["L1"].Value = "Advance Database Management System";
// Save Workbook

//Set Cell Value Dynamically
Random r = new ();
for (int i = 2; i <= 11; i++)
{
    sheet["A" + i].Value = r.Next(1, 100);
    sheet["B" + i].Value = r.Next(1, 100);
    sheet["C" + i].Value = r.Next(1, 100);
    sheet["D" + i].Value = r.Next(1, 100);
    sheet["E" + i].Value = r.Next(1, 100);
    sheet["F" + i].Value = r.Next(1, 100);
    sheet["G" + i].Value = r.Next(1, 100);
    sheet["H" + i].Value = r.Next(1, 100);
    sheet["I" + i].Value = r.Next(1, 100);
    sheet["J" + i].Value = r.Next(1, 100);
    sheet["K" + i].Value = r.Next(1, 100);
    sheet["L" + i].Value = r.Next(1, 100);
}
// Save Workbook

workbook.SaveAs("Result.xlsx");



//Application excel = new();
//Workbook wb = excel.Workbooks.Open(filePath);
//Worksheet ws = wb.Worksheets[1];
//object cell = ws.Cells[1, 1].value;


//Console.WriteLine(cell);
