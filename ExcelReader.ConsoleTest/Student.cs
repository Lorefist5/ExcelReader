using ExcelReader.Attributes;

class Student
{
    [Excel(Name = "Student name")]
    public string Name { get; set; } = default!;
    public int Age { get; set; }
    [Excel(DefaultValue = "No address")]
    public string Address { get; set; } 
};
