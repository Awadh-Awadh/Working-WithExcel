using System.Drawing;
using excelDemo;
using OfficeOpenXml;
using OfficeOpenXml.Style;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var file = new FileInfo(@"C:\Users\AwadhSaid\Desktop\Demos\learn.xlsx");

var people = GetSetUpData();
await SaveExcelFile(people, file);

static async Task SaveExcelFile(IEnumerable<PersonModel> people, FileInfo file)
{
    DeleteIfExists(file);
    
    using var package = new ExcelPackage(file);
    
    // create a worksheet
    var ws = package.Workbook.Worksheets.Add("MainReport");
   
    // add data to the sheet
    var range = ws.Cells["A2"].LoadFromCollection(people, true);
    range.AutoFitColumns();
    
    // formats the header row
    ws.Cells["A1"].Value = "Our cool report";
    ws.Cells["A1:C1"].Merge = true;
    ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    ws.Row(1).Style.Font.Size = 24;
    ws.Row(1).Style.Font.Color.SetColor(Color.Blue);

    ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    ws.Row(2).Style.Font.Bold = true;
    ws.Column(3).Width = 20;
    
    await package.SaveAsync();

}

static void DeleteIfExists(FileSystemInfo file)
{
    if (file.Exists)
    {
        file.Delete();
    }
}
static IEnumerable<PersonModel> GetSetUpData()
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