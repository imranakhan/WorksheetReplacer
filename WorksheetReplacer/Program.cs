//using Aspose.Cells;
using ClosedXML.Excel;

namespace WorksheetReplacer
{
    public class Program
    {
        static void Main(string[] args)
        {
            // 1: First argument is filename
            var fileName = args[0];

            // 2: it is column alphabet if column, or number if row
            var columnLetter = args[1];

            // 3: is previous string
            var previousString = args[2];

            // 4: new string to replace with
            var newString = args[3];

            // 5: Output file name
            var outputFile = args[4];

            using var workbook = new XLWorkbook(fileName);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            var cells = worksheet?.CellsUsed(x => x.WorksheetColumn().ColumnLetter() == columnLetter).ToList();
            var replaceCount = 0;

            foreach (IXLCell cell in cells)
            {
                if (cell.WorksheetColumn().ColumnLetter() == columnLetter && cell.GetValue<string>() == previousString)
                {
                    cell.Style = XLWorkbook.DefaultStyle;
                    cell.Value = newString;
                    replaceCount++;
                }
            }

            Console.WriteLine($"{replaceCount} instaces replaced with updated string");

            //worksheet.Clear(XLClearOptions.AllFormats);

            workbook.SaveAs(outputFile);
        }
    }
}
