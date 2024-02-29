using ExcelReader.Attributes;

namespace ExcelReader.ConsoleTest;

public class StudentDto
{
    [Excel(DefaultValue = "No address")]
    public string Address { get; set; }
    [Excel(IsProperty = false)]
    public bool IsAbove18 { get => Age > 18; }
    [Excel(Name = "Student name")]
    public string Name { get; set; } = default!;
    public int Age { get; set; }


}
