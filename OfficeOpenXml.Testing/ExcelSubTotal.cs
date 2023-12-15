namespace OfficeOpenXml.Testing;

/// <summary>
/// Class ExcelSubTotal.
/// </summary>
public class ExcelSubTotal
{
    /// <summary>
    /// Gets or sets the labels.
    /// </summary>
    /// <value>The labels.</value>
    public IDictionary<int, string> Labels { get; set; } = new Dictionary<int, string>();

    /// <summary>
    /// Gets or sets the column functions.
    /// </summary>
    /// <value>The column functions.</value>
    public IDictionary<int, int> ColumnFunctions { get; set; } = new Dictionary<int, int>();
}