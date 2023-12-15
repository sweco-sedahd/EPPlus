// See https://aka.ms/new-console-template for more information

using OfficeOpenXml;
using OfficeOpenXml.Testing;

Console.WriteLine("Hello, World!");

string fileLocation = "Files/AnchorPartnersPlus.Sales.Import.Template.xlsx";
//var fileLocation = "Files/OrdersHandback.xlsx";
var fileInfo = new FileInfo(fileLocation);
var stream = new MemoryStream(File.ReadAllBytes(fileLocation));

var package = new ExcelPackage(stream);
IDictionary<string, List<dynamic>> data = package.ToDictionary(2);

if (data.ContainsKey("CustomerSales") && data.TryGetValue("CustomerSales", out List<dynamic> dataTable))
{
    Console.WriteLine("Data Found");
}


Console.Read();