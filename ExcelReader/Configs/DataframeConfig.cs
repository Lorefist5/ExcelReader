using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelReader.Configs;

public class DataframeConfig
{
    public Color BackgroundColor { get; set; } = Color.White;
    public Color TextColor { get; set; } = Color.Black;
    public Color HeaderTextColor { get; set; } = Color.Black;
    public Color HeaderBackgroundColor { get; set; } = Color.White;
    public ExcelFillStyle FillStyle { get; set; } = ExcelFillStyle.Solid;
}
