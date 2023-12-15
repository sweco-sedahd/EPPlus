using OfficeOpenXml.Table;
using System.Data;

namespace OfficeOpenXml.Testing;

/// <summary>
/// Class ExcelExtensions.
/// </summary>
public static class ExcelExtensions
{
    /// <summary>
    /// Extracts a DataSet from the ExcelPackage.
    /// </summary>
    /// <param name="package">The Excel package.</param>
    /// <param name="firstRowContainsHeader">if set to <c>true</c> [first row contains header].</param>
    /// <returns>DataSet.</returns>
    public static DataSet ToDataSet(this ExcelPackage package, bool firstRowContainsHeader = false)
    {
        int headerRow = firstRowContainsHeader ? 1 : 0;
        return ToDataSet(package, headerRow);
    }

    /// <summary>
    /// Extracts a DataSet from the ExcelPackage.
    /// </summary>
    /// <param name="package">The Excel package.</param>
    /// <param name="headerRow">The header row. Use 0 if there is no header row. Value must be 0 or greater.</param>
    /// <returns>DataSet.</returns>
    /// <exception cref="ArgumentOutOfRangeException">headerRow - Must be 0 or greater.</exception>
    /// <exception cref="ArgumentException">headerRow must be 0 or greater.</exception>
    public static DataSet ToDataSet(this ExcelPackage package, int headerRow = 0)
    {
        if (headerRow < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(headerRow), headerRow, "Must be 0 or greater.");
        }

        var result = new DataSet();

        //todo First change required
        foreach (ExcelWorksheet sheet in package.Workbook.Worksheets.Where(w => w.Hidden == eWorkSheetHidden.Visible))
        {
            try
            {
                var table = new DataTable
                {
                    TableName = sheet.Name
                };

                int sheetStartRow = 1;

                if (headerRow > 1)
                {
                    sheetStartRow = headerRow;
                }

                IEnumerable<DataColumn> columns = from firstRowCell in sheet.Cells[sheetStartRow, 1, sheetStartRow, sheet.Dimension.End.Column]
                                                  select new DataColumn(headerRow > 0 ? firstRowCell.Value.ToString() : $"Column {firstRowCell.Start.Column}");

                table.Columns.AddRange(columns.ToArray());
                int startRow = headerRow > 0 ? sheetStartRow + 1 : sheetStartRow;

                for (int rowIndex = startRow; rowIndex <= sheet.Dimension.End.Row; rowIndex++)
                {
                    ExcelRange inputRow = sheet.Cells[rowIndex, 1, rowIndex, sheet.Dimension.End.Column];
                    DataRow dataRow = table.Rows.Add();

                    foreach (ExcelRangeBase cell in inputRow)
                    {
                        dataRow[cell.Start.Column - 1] = cell.Value;
                    }
                }

                result.Tables.Add(table);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        return result;
    }

    /// <summary>
    /// To the dictionary.
    /// </summary>
    /// <param name="package">The package.</param>
    /// <param name="headerRow">The header row.</param>
    /// <returns>IDictionary&lt;System.String, IList&lt;dynamic&gt;&gt;.</returns>
    /// <exception cref="ArgumentOutOfRangeException">headerRow - Must be 0 or greater.</exception>
    public static IDictionary<string, List<dynamic>> ToDictionary(this ExcelPackage package, int headerRow = 0)
    {
        if (headerRow < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(headerRow), headerRow, "Must be 0 or greater.");
        }

        var dictionary = new Dictionary<string, List<dynamic>>();

        foreach (ExcelWorksheet sheet in package.Workbook.Worksheets.Where(w => w.Hidden == eWorkSheetHidden.Visible))
        {
            try
            {
                string listName = sheet.Name;
                var excelData = new List<dynamic>();
                int sheetStartRow = 1;

                if (headerRow > 1)
                {
                    sheetStartRow = headerRow;
                }

                List<string> columns = sheet.Cells[sheetStartRow, 1, sheetStartRow, sheet.Dimension.End.Column].Select(c => c.Value.ToString()).ToList();
                int startRow = headerRow > 0 ? sheetStartRow + 1 : sheetStartRow;
                int columnCount = columns.Count;

                if (columnCount > 0)
                {
                    for (int rowIndex = startRow; rowIndex <= sheet.Dimension.End.Row; rowIndex++)
                    {
                        ExcelRange inputRow = sheet.Cells[rowIndex, 1, rowIndex, columnCount];
                        dynamic item = new DynamicDictionary();

                        for (int i = 0; i < columnCount; i++)
                        {
                            item.SetDynamicMember(columns[i], inputRow[rowIndex, i + 1].Value);
                        }

                        excelData.Add(item);
                    }

                    dictionary.Add(listName, excelData);
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        return dictionary;
    }

    /// <summary>
    /// Gets the tabl row functions.
    /// </summary>
    /// <param name="function">The function.</param>
    /// <returns>RowFunctions.</returns>
    public static RowFunctions ToRowFunctions(this int function)
    {
        return function switch
        {
            0 => RowFunctions.Average,
            1 => RowFunctions.Count,
            2 => RowFunctions.CountNums,
            3 => RowFunctions.Custom,
            4 => RowFunctions.Max,
            5 => RowFunctions.Min,
            6 => RowFunctions.None,
            7 => RowFunctions.StdDev,
            9 => RowFunctions.Var,
            _ => RowFunctions.Sum,
        };
    }
}