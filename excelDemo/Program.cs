using excelDemo;using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var file = new FileInfo(@"C:\Users\AwadhSaid\Desktop\Demos\learn.xlsx");

var people = GetSetUpData();

static List<PersonModel> GetSetUpData()
{
    List<PersonModel> output = new()
    {
        new() { Id = 1, FirstName = "Awadh", LastName = "Said" },
        new() { Id = 2, FirstName = "John", LastName = "Kamau" },
        new() { Id = 3, FirstName = "Dan", LastName = "Nakola" },
        new() { Id = 4, FirstName = "Elisha", LastName = "Misoi" }
    };
    return output;
}