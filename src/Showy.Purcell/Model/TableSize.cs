namespace Showy.Purcell;

/// <summary>
/// 表格尺寸，用于设置表格的行高和列宽，列宽分为最小列宽和最大列宽。<br /><br />
/// Table size, used to set the row height and column width of the table, and the column width is divided into the minimum column width and the maximum column width.
/// </summary>
public class TableSize
{
    /// <summary>
    /// 行高 Line height
    /// </summary>
    public decimal LineHeight { get; set; }

    /// <summary>
    /// 最小列宽 Minimum column width
    /// </summary>
    public decimal MinColumnWidth { get; set; }

    /// <summary>
    /// 最大列宽 Maximum column width
    /// </summary>
    public decimal MaxColumnWidth { get; set; }

    /// <summary>
    /// 字体大小 Font size
    /// </summary>
    public decimal FontSize { get; set; }

    /// <inheritdoc cref="TableSize"/>
    public TableSize()
    {
        LineHeight = TableSizes.Medium.LineHeight;
        MinColumnWidth = TableSizes.Medium.MinColumnWidth;
        MaxColumnWidth = TableSizes.Medium.MaxColumnWidth;
        FontSize = TableSizes.Medium.FontSize;
    }

    /// <inheritdoc cref="TableSize"/>
    /// <param name="lineHeight">行高 Line height</param>
    /// <param name="minColumnWidth">最小列宽 Minimum column width</param>
    /// <param name="maxColumnWidth">最大列宽 Maximum column width</param>
    /// <param name="fontSize">字体大小 Font size</param>
    public TableSize(decimal lineHeight, decimal minColumnWidth, decimal maxColumnWidth, decimal fontSize)
    {
        LineHeight = lineHeight;
        MinColumnWidth = minColumnWidth;
        MaxColumnWidth = maxColumnWidth;
        FontSize = fontSize;
    }
}