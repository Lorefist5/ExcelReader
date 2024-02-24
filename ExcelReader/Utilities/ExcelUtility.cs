using OfficeOpenXml;

namespace ExcelReader.Utilities;

public class ExcelUtility
{
    private readonly ExcelPackage _excelPackage;

    private ExcelPackage ExcelPackage => _excelPackage;

    public ExcelUtility(ExcelPackage excelPackage)
    {
        this._excelPackage = excelPackage;
    }
    public ExcelWorksheet GetExcelWorksheet(string sheetName)
    {
        return ExcelPackage.Workbook.Worksheets[GetSheetIndex(sheetName)];
    }
    public int GetSheetIndex(string sheetName)
    {
        for (int i = 0; i < ExcelPackage.Workbook.Worksheets.Count; i++)
        {
            if (ExcelPackage.Workbook.Worksheets[i].Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return 0;
    }
    public string GetExcelCoords(int row, int column)
    {
        string columnLetter = ((char)('A' + column - 1)).ToString();
        string cellCoords = $"{columnLetter}{row}";
        return cellCoords;
    }
    public List<string> GetColumnHeaders(ExcelWorksheet worksheet, int startRow, int startColumn)
    {
        List<string> excelColumnNames = new List<string>();
        int currentColumn = startColumn;

        while (worksheet.Cells[startRow, currentColumn].Value != null)
        {
            excelColumnNames.Add(worksheet.Cells[startRow, currentColumn].Value.ToString()!);
            currentColumn++;
        }

        return excelColumnNames;
    }
}
