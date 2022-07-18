using System.Drawing;
using excelDemo;
using OfficeOpenXml;
using OfficeOpenXml.Style;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var file = new FileInfo(@"C:\Users\AwadhSaid\Desktop\Demos\learn.xlsx");

var people = GetSetUpData();
await SaveExcelFile(people, file);
var peopleFromExcel = await LoadFromExcel(file);

foreach (var p in peopleFromExcel)
{
    Console.WriteLine($"{p.Id} {p.FirstName}, {p.LastName}");
}

static async Task<List<PersonModel>> LoadFromExcel(FileInfo file)
{
    List<PersonModel> output = new() { };
    using var package = new ExcelPackage(file);
    
    await package.LoadAsync(file);
    var ws = package.Workbook.Worksheets[0];
    // first two rows are for headers
    var row = 3;
    var column = 1;
    while(string.IsNullOrWhiteSpace(ws.Cells[row,column].Value?.ToString()) == false)
    {
        PersonModel person = new();
        person.Id = int.Parse(ws.Cells[row, column].Value.ToString());
        person.FirstName = ws.Cells[row, column = 1].ToString();
        person.LastName = ws.Cells[row, column + 2].ToString();
        output.Add(person);
        row+=1;
    }

    return output;

}

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