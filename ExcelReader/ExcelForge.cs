using ExcelReader.Attributes;
using ExcelReader.Configs;
using ExcelReader.Utilities;
using OfficeOpenXml;
using System.Reflection;

namespace ExcelReader;

public class ExcelForge
{
    private ExcelPackage? _excelPackage;
    private ExcelPackage ExcelPackage
    {
        get
        {
            EnsureExcelPackageCreated();
            return _excelPackage!;
        }
        set => _excelPackage = value;
    }
    private ExcelUtility _excelUtility = default!;
    private readonly DataframeConfig _dataframeConfig;

    public ExcelForge(DataframeConfig? dataframeConfig = null)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        _dataframeConfig = dataframeConfig ??= new DataframeConfig();
    }

    public ExcelForge CreateExcelPackage()
    {
        ExcelPackage excelPackage = new ExcelPackage();
        ExcelPackage = excelPackage;
        _excelUtility = new ExcelUtility(ExcelPackage);
        return this;
    }
    public ExcelForge CreateExcelSheet(string sheetName)
    {
        EnsureExcelPackageCreated();
        ExcelPackage.Workbook.Worksheets.Add(sheetName);
        return this;
    }
    public ExcelForge ReadExcelFile(string filePath)
    {
        if (!Path.HasExtension(filePath)) filePath += ".xlsx";
        if (!Path.Exists(filePath)) throw new FileNotFoundException();


        ExcelPackage excelPackage = new ExcelPackage(filePath);
        ExcelPackage = excelPackage;
        _excelUtility = new ExcelUtility(ExcelPackage);
        return this;
    }
    public ExcelForge AddParagraph(string paragraphName, int row, int column, string sheetName = "Sheet1")
    {
        EnsureExcelPackageCreated();


        var worksheet = ExcelPackage.Workbook.Worksheets[_excelUtility.GetSheetIndex(sheetName)];

        worksheet.Cells[row, column].Value = paragraphName;


        return this;
    }
    public List<T> ReadDataframe<T>(int startRow, int startColumn, string sheetName = "Sheet1") where T : class, new()
    {
        EnsureExcelPackageCreated();

        var worksheet = ExcelPackage.Workbook.Worksheets[_excelUtility.GetSheetIndex(sheetName)];

        var properties = typeof(T).GetProperties().ToList();

        List<T> result = new List<T>();
        List<string> excelColumnNames = _excelUtility.GetColumnHeaders(worksheet, startRow, startColumn);

        int currentRow = startRow + 1; // Move to the next row after reading column headers

        while (worksheet.Cells[currentRow, startColumn].Value != null)
        {
            WorksheetUtility worksheetUtility = new WorksheetUtility(worksheet,excelColumnNames, startColumn, startRow);
            T item = PopulateObjectFromRow<T>(worksheetUtility, properties, excelColumnNames, currentRow, startColumn);
            result.Add(item);
            currentRow++;
        }

        return result;
    }



    public ExcelForge AddDataframe<T>(List<T> data, int startRow, int startColumn, string sheetName = "Sheet1") where T : class
    {
        EnsureExcelPackageCreated();
        var worksheet = _excelUtility.GetExcelWorksheet(sheetName);

        var properties = typeof(T).GetProperties().ToList();

        int headerColumn = startColumn;

        foreach (var property in properties)
        {
            var excelAttributes = ExcelAttribute.GetDefaultExcelAttribute(property);

            if (excelAttributes.IsProperty)
            {
                string columnName = GetColumnName(property);
                string coordinates = _excelUtility.GetExcelCoords(startRow, headerColumn);
                ExcelRange cell = worksheet.Cells[coordinates];
                cell.Value = columnName;
                cell.Style.Font.Color.SetColor(_dataframeConfig.HeaderTextColor);
                cell.Style.Fill.PatternType = _dataframeConfig.FillStyle;
                cell.Style.Fill.BackgroundColor.SetColor(_dataframeConfig.HeaderBackgroundColor);
                headerColumn++;
            }
        }

        int currentRow = startRow + 1;

        foreach (var item in data)
        {
            int currentColumn = startColumn;

            foreach (var property in properties)
            {
                var excelAttributes = ExcelAttribute.GetDefaultExcelAttribute(property);
                if(excelAttributes.IsProperty)
                {
                    string coordinates = _excelUtility.GetExcelCoords(currentRow, currentColumn);

                    var cellValue = property.GetValue(item) ?? excelAttributes.DefaultValue;

                    ExcelRange cell = worksheet.Cells[coordinates];
                    if (property.PropertyType == typeof(int))
                    {
                        cell.Value = cellValue;
                    }
                    else
                    {
                        cell.Value = cellValue?.ToString();
                    }
                    cell.Style.Font.Color.SetColor(_dataframeConfig.TextColor);
                    cell.Style.Fill.PatternType = _dataframeConfig.FillStyle;
                    cell.Style.Fill.BackgroundColor.SetColor(_dataframeConfig.BackgroundColor);
                    currentColumn++;
                }
            }

            currentRow++;
        }

        return this;
    }
    public ExcelForge SaveAs(string filePath)
    {
        EnsureExcelPackageCreated();
        if (!Path.HasExtension(filePath)) filePath += ".xlsx";
        ExcelPackage.SaveAs(new FileInfo(filePath));
        return this;
    }
    public ExcelForge SaveChanges()
    {
        EnsureExcelPackageCreated();

        ExcelPackage.Save();
        return this;
    }
    private string GetColumnName(PropertyInfo property)
    {
        var excelAttribute = property.GetCustomAttribute<ExcelAttribute>();
        return excelAttribute?.Name ?? property.Name;
    }
    private T PopulateObjectFromRow<T>(WorksheetUtility worksheetUtility, List<PropertyInfo> properties, List<string> excelColumnNames, int currentRow, int startColumn) where T : class, new()
    {
        T item = new T();
        int currentColumn = startColumn;
        var filteredProperties = properties.Where(p => p.GetCustomAttribute<ExcelAttribute>() == null || p.GetCustomAttribute<ExcelAttribute>().IsProperty).ToList();
        foreach (var property in filteredProperties)
        {
            var excelAttribute = ExcelAttribute.GetDefaultExcelAttribute(property);

            string columnName = excelAttribute.Name ?? property.Name;

            int matchingColumnIndex = worksheetUtility.GetMatchingColumnIndex(columnName);

            if (matchingColumnIndex > 0)
            {
                var convertedValue = worksheetUtility.GetObjectFromCell(property, currentRow, matchingColumnIndex);
                property.SetValue(item, convertedValue);
            }


            currentColumn++;
        }

        return item;
    }
    private void EnsureExcelPackageCreated()
    {
        if (_excelPackage == null)
        {
            throw new InvalidOperationException("Excel file not created. Call CreateExcelFile first or ReadExcelFile.");
        }
    }
}
