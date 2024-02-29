using ExcelReader;
using ExcelReader.Configs;
using ExcelReader.ConsoleTest;
using System.Drawing;

DataframeConfig dataframeConfig = new DataframeConfig()
{
    BackgroundColor = Color.Gray,
    TextColor = Color.White,
    HeaderBackgroundColor = Color.Black,
    HeaderTextColor = Color.White,
};
ExcelForge excelForge = new ExcelForge(dataframeConfig);

List<Student> students = new List<Student>()
{
    new Student { Name = "John Doe", Age = 20, Address = "123 Main St" },
    new Student { Name = "Jane Smith", Age = 22, Address = "456 Oak Ave" },
    new Student { Name = "Bob Johnson", Age = 21, Address = "789 Pine Blvd" },
    new Student { Name = "Alice Brown", Age = 19, Address = "101 Cedar Ln" },
    new Student { Name = "Charlie Davis", Age = 23, Address = "202 Elm Rd" },
    new Student { Name = "Eva White", Age = 18, Address = "303 Birch Dr" },
    new Student { Name = "Frank Miller", Age = 24, Address = "404 Maple Ct" },
    new Student { Name = "Grace Taylor", Age = 20, Address = "505 Walnut Pl" },
    new Student { Name = "David Clark", Age = 22, Address = "606 Spruce Ter" },
    new Student { Name = "Sophie Turner", Age = 21 }
};



excelForge.CreateExcelPackage().CreateExcelSheet("Sheet1").AddDataframe(students,1,1).SaveAs("Students");


var studentsInExcel = excelForge.ReadExcelFile("Students").ReadDataframe<StudentDto>(1, 1);

var filteredStudents = studentsInExcel.Where(s => s.IsAbove18).ToList();


foreach(var filteredStduent in filteredStudents)
{
    Console.WriteLine(filteredStduent.Name);
}

Console.ReadLine();