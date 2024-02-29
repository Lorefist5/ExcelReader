using ExcelReader.Attributes;
using OfficeOpenXml;
using System.Reflection;

namespace ExcelReader.Utilities;

public class WorksheetUtility
{
    private readonly ExcelWorksheet _excelWorksheet;
    private readonly List<string> _excelColumnNames;
    private int _startColumn;
    private int _startRow;
    public WorksheetUtility(ExcelWorksheet excelWorksheet,List<string> excelColumnNames, int startColumn, int startRow)
    {
        this._excelWorksheet = excelWorksheet;
        this._excelColumnNames = excelColumnNames;
        this._startColumn = startColumn;
        this._startRow = startRow;
    }
    public int GetMatchingColumnIndex(string columnName)
    {
        var matchingColumnName = _excelColumnNames.FirstOrDefault(c => c.Equals(columnName, StringComparison.OrdinalIgnoreCase));
        if (matchingColumnName != null)
        {
            return _excelColumnNames.IndexOf(matchingColumnName) + _startColumn;
        }
        return -1;
    }
    public object? GetObjectFromCell(PropertyInfo property, int currentRow, int propertyColumnIndex)
    {
        var excelAttribute = ExcelAttribute.GetDefaultExcelAttribute(property);

        var cellValue = _excelWorksheet.Cells[currentRow, propertyColumnIndex].Value?.ToString();
        if (cellValue != null)
        {
            var convertedValue = Convert.ChangeType(cellValue, excelAttribute.DataType ?? typeof(string));
            return convertedValue;
        }
        
        return null;
    }

}
