namespace ExcelReader.Attributes;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
public class ExcelAttribute : Attribute
{
    public ExcelAttribute(string? name = null, bool isProperty = true, object? defaultValue = null, Type? dataType = null)
    {
        Name = name;
        IsProperty = isProperty;
        DataType = dataType;
        DefaultValue = defaultValue;
    }

    public string? Name { get; set; }
    public object? DefaultValue { get; set; }
    public bool IsProperty { get; set; }
    public Type? DataType { get; set; }
}


