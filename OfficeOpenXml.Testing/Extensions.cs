using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using System.ComponentModel;
using System.Data;
using Humanizer;
using System.Text;

namespace OfficeOpenXml.Testing;

public static class Extensions
{
    /// <summary>
    /// Writes the excel data.
    /// </summary>
    /// <param name="dataTable">The data table.</param>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="excelSubTotal">The excel sub total.</param>
    /// <returns>System.Byte[].</returns>
    public static byte[] ToExcel(this DataTable dataTable, string sheetName, ExcelSubTotal excelSubTotal = null)
    {
        var excelPackage = new ExcelPackage();
        ExcelNamedStyleXml decimalStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableDecimal");
        ExcelNamedStyleXml numberStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableNumber");
        ExcelNamedStyleXml dateStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableDate");
        ExcelNamedStyleXml timeStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableTime");
        decimalStyle.Style.Numberformat.Format = "#,##0.00";
        numberStyle.Style.Numberformat.Format = "#,##0";
        dateStyle.Style.Numberformat.Format = @"dd\/mm\/yyyy";
        timeStyle.Style.Numberformat.Format = "HH:mm";
        WriteToExcel(dataTable, sheetName, excelSubTotal, excelPackage);
        return excelPackage.GetAsByteArray();
    }

    /// <summary>
    /// indicates whether this string is null, empty, or consists only of white-space characters.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns><c>true</c> if [is null or white space] [the specified value]; otherwise, <c>false</c>.</returns>
    public static bool IsNullOrWhiteSpace(this string value)
    {
        return string.IsNullOrWhiteSpace(value);
    }

    /// <summary>
    /// Determines whether [is not null or white space] [the specified value].
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns><c>true</c> if [is not null or white space] [the specified value]; otherwise, <c>false</c>.</returns>
    public static bool IsNotNullOrWhiteSpace(this string value)
    {
        return !string.IsNullOrWhiteSpace(value);
    }

    /// <summary>
    /// Converts a camel string to a displayable string.
    /// </summary>
    /// <param name="value">The camel string</param>
    /// <returns>A string with all capitals preceded by blanks</returns>
    public static string ToTitleCase(this string value)
    {
        return value.IsNullOrWhiteSpace() ? value : value.Titleize();
    }

    /// <summary>
    /// Cleans the white space.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns>System.String.</returns>
    public static string CleanWhiteSpace(this string value)
    {
        var cleanString = new StringBuilder(1024);

        foreach (char character in value.Where(character => !char.IsWhiteSpace(character)))
        {
            cleanString.Append(character);
        }

        return cleanString.ToString();
    }

    /// <summary>
    /// Writes to excel.
    /// </summary>
    /// <param name="dataTable">The data table.</param>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="excelSubTotal">The excel sub total.</param>
    /// <param name="excelPackage">The excel package.</param>
    internal static void WriteToExcel(this DataTable dataTable, string sheetName, ExcelSubTotal excelSubTotal, ExcelPackage excelPackage)
    {
        ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add(sheetName);
        workSheet.View.ShowGridLines = false;
        int row = 1;
        int column = 1;

        //write the headers
        for (int i = 0; i < dataTable.Columns.Count; i++)
        {
            workSheet.Cells[row, column++].Value = dataTable.Columns[i].ColumnName.ToTitleCase();
        }

        row++;

        foreach (DataRow dataRow in dataTable.Rows)
        {
            column = 1;

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                try
                {
                    //workSheet.Cells[row, column++].Value = dataRow[i];
                    object value = dataRow[i];
                    workSheet.Cells[row, column++].Value = value;

                    if (value != DBNull.Value && value is string && value.ToString()?.Length >= 256)
                    {
                        workSheet.Cells[row, column].Style.WrapText = true;
                    }
                }
                catch (Exception)
                {
                    workSheet.Cells[row, column].Value = null;
                }
            }

            row++;
        }

        try
        {
            ExcelRange range = workSheet.Cells[1, 1, row - 1, column - 1];
            ExcelTable table = workSheet.Tables.Add(range, sheetName.CleanWhiteSpace());
            table.ShowHeader = row > 2;
            table.TableStyle = TableStyles.Medium9;

            if (excelSubTotal?.ColumnFunctions != null)
            {
                table.ShowTotal = true;

                if (excelSubTotal.Labels != null)
                {
                    foreach ((int key, string value) in excelSubTotal.Labels)
                    {
                        table.Columns[key].TotalsRowLabel = value;
                    }
                }

                foreach ((int key, int value) in excelSubTotal.ColumnFunctions)
                {
                    table.Columns[key].TotalsRowFunction = value.ToRowFunctions();
                }
            }
        }
        catch (Exception exception)
        {
            //Console.WriteLine(exception);
        }

        int formatColumn = 1;

        for (int i = 0; i < dataTable.Columns.Count; i++)
        {
            Type dataType = dataTable.Columns[i].DataType;

            if (row > 2 && (dataType == typeof(DateTime?) || dataType == typeof(DateTime) || dataType == typeof(DateOnly?) || dataType == typeof(DateOnly)))
            {
                workSheet.Cells[2, formatColumn, row - 1, formatColumn].StyleName = "TableDate";
            }

            if (row > 2 && (dataType == typeof(TimeOnly?) || dataType == typeof(TimeOnly)))
            {
                workSheet.Cells[2, formatColumn, row - 1, formatColumn].StyleName = "TableTime";
            }

            if (row > 2 && (dataType == typeof(decimal?) || dataType == typeof(decimal) || dataType == typeof(double?) || dataType == typeof(double) || dataType == typeof(float?) || dataType == typeof(float)))
            {
                workSheet.Cells[2, formatColumn, row - 1, formatColumn].StyleName = "TableDecimal";
            }

            formatColumn++;
        }

        // If sheet dimensions are up to 10,000 cells then autofit otherwise only autofit the header row.
        if (workSheet.Dimension.Rows * workSheet.Dimension.Columns <= 10000)
        {
            workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
        }
        else
        {
            workSheet.Cells[1, 1, 1, column - 1].AutoFitColumns();
        }
    }

    /// <summary>
    /// To the dynamic excel.
    /// </summary>
    /// <param name="source">The source.</param>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="columnHeaders">The column headers.</param>
    /// <param name="excelSubTotal">The excel sub total.</param>
    /// <returns>System.Byte[].</returns>
    public static byte[] ToDynamicExcel(this IEnumerable<dynamic> source, string sheetName, IList<string> columnHeaders, ExcelSubTotal excelSubTotal = null)
    {
        var excelPackage = new ExcelPackage();
        ExcelNamedStyleXml decimalStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableDecimal");
        ExcelNamedStyleXml numberStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableNumber");
        ExcelNamedStyleXml dateStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableDate");
        ExcelNamedStyleXml timeStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableTime");
        decimalStyle.Style.Numberformat.Format = "#,##0.00";
        numberStyle.Style.Numberformat.Format = "#,##0";
        dateStyle.Style.Numberformat.Format = @"dd\/mm\/yyyy";
        timeStyle.Style.Numberformat.Format = "HH:mm";
        source.WriteToExcel(sheetName, columnHeaders, excelSubTotal, excelPackage);
        return excelPackage.GetAsByteArray();
    }

    /// <summary>
    /// Writes to excel.
    /// </summary>
    /// <param name="source">The source.</param>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="columnHeaders">The column headers.</param>
    /// <param name="excelSubTotal">The excel sub total.</param>
    /// <param name="excelPackage">The excel package.</param>
    internal static void WriteToExcel(this IEnumerable<dynamic> source, string sheetName, IList<string> columnHeaders, ExcelSubTotal excelSubTotal, ExcelPackage excelPackage)
    {
        ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add(sheetName);
        workSheet.View.ShowGridLines = false;
        int row = 1;
        int column = 1;
        var columnTypes = new Dictionary<string, Type>();

        if (columnHeaders?.Count > 0)
        {
            //write the headers
            foreach (string columnName in columnHeaders)
            {
                workSheet.Cells[row, column++].Value = columnName;
            }

            row++;

            if (source != null)
            {
                foreach (dynamic item in source)
                {
                    column = 1;

                    foreach (string columnName in columnHeaders)
                    {
                        try
                        {
                            object value = item.GetMemberValue(columnName);
                            workSheet.Cells[row, column++].Value = value;

                            if (value != null && value != DBNull.Value)
                            {
                                columnTypes.TryAdd(columnName, value.GetType());

                                if (value is string)
                                {
                                    string stringValue = value.ToString();

                                    if (!stringValue.IsNullOrWhiteSpace() && (stringValue?.Length >= 256 || (stringValue?.Length >= 100 && (stringValue.Contains(Environment.NewLine) || stringValue.Contains((char)13)))))
                                    {
                                        workSheet.Cells[row, column].Style.WrapText = true;
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            workSheet.Cells[row, column].Value = null;
                        }
                    }

                    row++;
                }
            }

            try
            {
                ExcelRange range = workSheet.Cells[1, 1, row - 1, column - 1];
                ExcelTable table = workSheet.Tables.Add(range, sheetName.CleanWhiteSpace());
                table.ShowHeader = row > 2;
                table.TableStyle = TableStyles.Medium9;

                if (excelSubTotal?.ColumnFunctions != null)
                {
                    table.ShowTotal = true;

                    if (excelSubTotal.Labels != null)
                    {
                        foreach ((int key, string value) in excelSubTotal.Labels)
                        {
                            table.Columns[key].TotalsRowLabel = value;
                        }
                    }

                    foreach ((int key, int value) in excelSubTotal.ColumnFunctions)
                    {
                        table.Columns[key].TotalsRowFunction = value.ToRowFunctions();
                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }

            int formatColumn = 1;

            foreach (string columnName in columnHeaders)
            {
                try
                {
                    foreach ((string key, Type value) in columnTypes)
                    {
                        if (key == columnName)
                        {
                            if (row > 2 && (value == typeof(DateTime?) || value == typeof(DateTime) || value == typeof(DateOnly?) || value == typeof(DateOnly)))
                            {
                                workSheet.Cells[2, formatColumn, row - 1, formatColumn].StyleName = "TableDate";
                            }

                            if (row > 2 && (value == typeof(TimeOnly?) || value == typeof(TimeOnly)))
                            {
                                workSheet.Cells[2, formatColumn, row - 1, formatColumn].StyleName = "TableTime";
                            }

                            if (row > 2 && (value == typeof(decimal?) || value == typeof(decimal) || value == typeof(double?) || value == typeof(double) || value == typeof(float?) || value == typeof(float)))
                            {
                                workSheet.Cells[2, formatColumn, row - 1, formatColumn].StyleName = "TableDecimal";
                            }

                            break;
                        }
                    }
                }
                catch (Exception)
                {
                    workSheet.Cells[row, column].Value = null;
                }

                formatColumn++;
            }

            // If sheet dimensions are up to 10,000 cells then autofit otherwise only autofit the header row.
            if (workSheet.Dimension.Rows * workSheet.Dimension.Columns <= 10000)
            {
                workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
            }
            else
            {
                workSheet.Cells[1, 1, 1, column - 1].AutoFitColumns();
            }
        }
    }

    /// <summary>
    /// To the excel.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="source">The source.</param>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="excludeColumns">The exclude columns.</param>
    /// <param name="columnHeaders">The column headers.</param>
    /// <param name="excelSubTotal">The excel sub total.</param>
    /// <returns>System.Byte[].</returns>
    public static byte[] ToExcel<T>(this IEnumerable<T> source, string sheetName, IList<string> excludeColumns = null, IList<string> columnHeaders = null, ExcelSubTotal excelSubTotal = null)
    {
        var excelPackage = new ExcelPackage();
        ExcelNamedStyleXml decimalStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableDecimal");
        ExcelNamedStyleXml numberStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableNumber");
        ExcelNamedStyleXml dateStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableDate");
        ExcelNamedStyleXml timeStyle = excelPackage.Workbook.Styles.CreateNamedStyle("TableTime");
        decimalStyle.Style.Numberformat.Format = "#,##0.00";
        numberStyle.Style.Numberformat.Format = "#,##0";
        dateStyle.Style.Numberformat.Format = @"dd\/mm\/yyyy";
        timeStyle.Style.Numberformat.Format = "HH:mm";
        source.WriteToExcel(sheetName, excludeColumns, columnHeaders, excelSubTotal, excelPackage);
        return excelPackage.GetAsByteArray();
    }

    /// <summary>
    /// Writes to excel.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="source">The source.</param>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="excludeColumns">The exclude columns.</param>
    /// <param name="columnHeaders">The column headers.</param>
    /// <param name="excelSubTotal">The excel sub total.</param>
    /// <param name="excelPackage">The excel package.</param>
    internal static void WriteToExcel<T>(this IEnumerable<T> source, string sheetName, IList<string> excludeColumns, IList<string> columnHeaders, ExcelSubTotal excelSubTotal, ExcelPackage excelPackage)
    {
        ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add(sheetName);
        workSheet.View.ShowGridLines = false;

        Type type = typeof(T);
        PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(type);
        int row = 1;
        int column = 1;
        excludeColumns ??= new List<string>();

        //write the headers
        if (columnHeaders != null)
        {
            foreach (string columnHeader in columnHeaders)
            {
                workSheet.Cells[row, column++].Value = columnHeader;
            }
        }
        else
        {
            foreach (PropertyDescriptor property in properties)
            {
                if (!property.PropertyType.CanUsePropertyType())
                {
                    continue; //shallow only
                }

                if (excludeColumns.Contains(property.Name))
                {
                    continue;
                }

                workSheet.Cells[row, column++].Value = property.Name;
            }
        }

        row++;

        //write the data
        foreach (T item in source)
        {
            column = 1;

            foreach (PropertyDescriptor property in properties)
            {
                if (!property.PropertyType.CanUsePropertyType())
                {
                    continue; //shallow only
                }

                if (!excludeColumns.Contains(property.Name))
                {
                    try
                    {
                        object value = property.GetValue(item);
                        workSheet.Cells[row, column].Value = value;

                        if (value != null && value != DBNull.Value && value is string)
                        {
                            string stringValue = value.ToString();

                            if (!stringValue.IsNullOrWhiteSpace() && (stringValue?.Length >= 256 || (stringValue?.Length >= 100 && (stringValue.Contains(Environment.NewLine) || stringValue.Contains((char)13)))))
                            {
                                workSheet.Cells[row, column].Style.WrapText = true;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        workSheet.Cells[row, column].Value = null;
                    }
                    finally
                    {
                        column++;
                    }
                }
            }

            row++;
        }

        try
        {
            ExcelRange range = workSheet.Cells[1, 1, row - 1, column - 1];
            ExcelTable table = workSheet.Tables.Add(range, sheetName.CleanWhiteSpace());
            table.ShowHeader = row > 2;
            table.TableStyle = TableStyles.Medium9;

            if (excelSubTotal?.ColumnFunctions != null)
            {
                table.ShowTotal = true;

                if (excelSubTotal.Labels != null)
                {
                    foreach ((int key, string value) in excelSubTotal.Labels)
                    {
                        table.Columns[key].TotalsRowLabel = value;
                    }
                }

                foreach ((int key, int value) in excelSubTotal.ColumnFunctions)
                {
                    table.Columns[key].TotalsRowFunction = value.ToRowFunctions();
                }
            }
        }
        catch (Exception exception)
        {
            workSheet.Cells[2, 2].Value = exception;
        }

        int formatColumn = 1;

        foreach (PropertyDescriptor property in properties)
        {
            if (!property.PropertyType.CanUsePropertyType())
            {
                continue; //shallow only
            }

            if (excludeColumns.Contains(property.Name))
            {
                continue;
            }

            if (row > 2 && (property.PropertyType == typeof(DateTime?) || property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateOnly?) || property.PropertyType == typeof(DateOnly)))
            {
                workSheet.Cells[2, formatColumn, row - 1, formatColumn].StyleName = "TableDate";
            }

            if (row > 2 && (property.PropertyType == typeof(TimeOnly?) || property.PropertyType == typeof(TimeOnly)))
            {
                workSheet.Cells[2, formatColumn, row - 1, formatColumn].StyleName = "TableTime";
            }

            if (row > 2 && (property.PropertyType == typeof(decimal?) || property.PropertyType == typeof(decimal) || property.PropertyType == typeof(double?) || property.PropertyType == typeof(double) || property.PropertyType == typeof(float?) || property.PropertyType == typeof(float)))
            {
                workSheet.Cells[2, formatColumn, row - 1, formatColumn].StyleName = "TableDecimal";
            }

            formatColumn++;
        }

        // If sheet dimensions are up to 10,000 cells then autofit otherwise only autofit the header row.
        if (workSheet.Dimension.Rows * workSheet.Dimension.Columns <= 10000)
        {
            workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
        }
        else
        {
            workSheet.Cells[1, 1, 1, column - 1].AutoFitColumns();
        }
    }

    /// <summary>
    /// Determines whether this instance [can use property type] the specified type.
    /// </summary>
    /// <param name="type">The type.</param>
    /// <returns><c>true</c> if this instance [can use property type] the specified type; otherwise, <c>false</c>.</returns>
    public static bool CanUsePropertyType(this Type type)
    {
        //only strings and value types
        return !type.IsArray && (type.IsValueType || type == typeof(string));
    }
}