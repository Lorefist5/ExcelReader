using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelReader.Models;

public class DataframeConfig
{
    public Color BackgroundColor { get; set; } = Color.White;
    public Color TextColor { get; set; } = Color.Black;
    public Color HeaderTextColors { get; set; } = Color.Black;
    public Color HeaderBackgroundColor { get; set; } = Color.White;
    public ExcelFillStyle FillStyle { get; set; }  = ExcelFillStyle.Solid;
}
