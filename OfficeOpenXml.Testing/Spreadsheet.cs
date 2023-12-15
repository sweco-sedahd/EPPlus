using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Style;
using System.Data;
using System.Drawing;

namespace OfficeOpenXml.Testing;

/// <summary>
/// Class Spreadsheet.
/// </summary>
public class Spreadsheet : IDisposable
{
    private readonly ExcelPackage _excelPackage;
    private bool _disposed;

    /// <summary>
    /// Initializes a new instance of the <see cref="Spreadsheet"/> class.
    /// </summary>
    public Spreadsheet()
    {
        _excelPackage = new ExcelPackage();
        ExcelNamedStyleXml decimalStyle = _excelPackage.Workbook.Styles.CreateNamedStyle("TableDecimal");
        ExcelNamedStyleXml numberStyle = _excelPackage.Workbook.Styles.CreateNamedStyle("TableNumber");
        ExcelNamedStyleXml dateStyle = _excelPackage.Workbook.Styles.CreateNamedStyle("TableDate");
        ExcelNamedStyleXml timeStyle = _excelPackage.Workbook.Styles.CreateNamedStyle("TableTime");
        decimalStyle.Style.Numberformat.Format = "#,##0.00";
        numberStyle.Style.Numberformat.Format = "#,##0";
        dateStyle.Style.Numberformat.Format = @"dd\/mm\/yyyy";
        timeStyle.Style.Numberformat.Format = "HH:mm";

        ExcelNamedStyleXml goodStyle = _excelPackage.Workbook.Styles.CreateNamedStyle("CellGood");
        goodStyle.Style.Font.Name = "Calibri";
        goodStyle.Style.Font.Family = 2;
        goodStyle.Style.Font.Size = 11;
        goodStyle.Style.Font.Color.SetColor(ColorTranslator.FromHtml("#006100"));
        goodStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
        goodStyle.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C6EFCE"));

        ExcelNamedStyleXml badStyle = _excelPackage.Workbook.Styles.CreateNamedStyle("CellBad");
        badStyle.Style.Font.Name = "Calibri";
        badStyle.Style.Font.Family = 2;
        badStyle.Style.Font.Size = 11;
        badStyle.Style.Font.Color.SetColor(ColorTranslator.FromHtml("#9C0006"));
        badStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
        badStyle.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFC7CE"));

        ExcelNamedStyleXml neutralStyle = _excelPackage.Workbook.Styles.CreateNamedStyle("CellNeutral");
        neutralStyle.Style.Font.Name = "Calibri";
        neutralStyle.Style.Font.Family = 2;
        neutralStyle.Style.Font.Size = 11;
        neutralStyle.Style.Font.Color.SetColor(ColorTranslator.FromHtml("#9C5700"));
        neutralStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
        neutralStyle.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFEB9C"));
    }

    /// <summary>
    /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// To the spreadsheet.
    /// </summary>
    /// <returns>System.Byte[].</returns>
    public byte[] ToSpreadsheet()
    {
        return _excelPackage.GetAsByteArray();
    }

    /// <summary>
    /// Gets the worksheet.
    /// </summary>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <returns>ExcelWorksheet.</returns>
    public ExcelWorksheet GetWorksheet(string sheetName)
    {
        return _excelPackage.Workbook.Worksheets[sheetName];
    }

    /// <summary>
    /// Writes the worksheet.
    /// </summary>
    /// <param name="dataTable">The data table.</param>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="excelSubTotal">The excel sub total.</param>
    public void WriteWorksheet(DataTable dataTable, string sheetName, ExcelSubTotal excelSubTotal = null)
    {
        dataTable.WriteToExcel(sheetName, excelSubTotal, _excelPackage);
    }

    /// <summary>
    /// Writes the worksheet.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="source">The source.</param>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="excludeColumns">The exclude columns.</param>
    /// <param name="columnHeaders">The column headers.</param>
    /// <param name="excelSubTotal">The excel sub total.</param>
    public void WriteWorksheet<T>(IEnumerable<T> source, string sheetName, IList<string> excludeColumns = null, IList<string> columnHeaders = null, ExcelSubTotal excelSubTotal = null)
    {
        source.WriteToExcel(sheetName, excludeColumns, columnHeaders, excelSubTotal, _excelPackage);
    }

    /// <summary>
    /// Writes the worksheet.
    /// </summary>
    /// <param name="source">The source.</param>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="columnHeaders">The column headers.</param>
    /// <param name="excelSubTotal">The excel sub total.</param>
    public void WriteWorksheet(IEnumerable<dynamic> source, string sheetName, IList<string> columnHeaders, ExcelSubTotal excelSubTotal = null)
    {
        source.WriteToExcel(sheetName, columnHeaders, excelSubTotal, _excelPackage);
    }

    /// <summary>
    /// Releases unmanaged and - optionally - managed resources.
    /// </summary>
    /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
    private void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                _excelPackage?.Dispose();
            }

            _disposed = true;
        }
    }

    /// <summary>
    /// Finalizes this instance.
    /// </summary>
    ~Spreadsheet()
    {
        Dispose(false);
    }
}