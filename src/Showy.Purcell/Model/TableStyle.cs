using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace Showy.Purcell;

/// <summary>
/// 表格样式 Table style
/// </summary>
public class TableStyle
{
    /// <inheritdoc cref="TableStyle"/>
    public TableStyle()
    {
        HeaderBgColor = TableStyles.Default.HeaderBgColor;
        HeaderTextColor = TableStyles.Default.HeaderTextColor;
        LineHeight = TableSizes.Medium.LineHeight;
        MinColumnWidth = TableSizes.Medium.MinColumnWidth;
        MaxColumnWidth = TableSizes.Medium.MaxColumnWidth;
        FontSize = TableSizes.Medium.FontSize;
    }

    /// <inheritdoc cref="TableStyle"/>
    /// <param name="headerBgColor">表头背景色 Header background color</param>
    /// <param name="headerTextColor">表头文本色 Header text color</param>
    /// <param name="lineHeight">行高 Line height</param>
    /// <param name="minColumnWidth">最小列宽 Minimum column width</param>
    /// <param name="maxColumnWidth">最大列宽 Maximum column width</param>
    /// <param name="fontSize">字体大小 Font size</param>
    public TableStyle(Color headerBgColor, Color headerTextColor, decimal lineHeight, decimal minColumnWidth, decimal maxColumnWidth, decimal fontSize)
    {
        HeaderBgColor = headerBgColor;
        HeaderTextColor = headerTextColor;
        LineHeight = lineHeight;
        MinColumnWidth = minColumnWidth;
        MaxColumnWidth = maxColumnWidth;
        FontSize = fontSize;
    }

    /// <inheritdoc cref="TableStyle"/>
    /// <param name="headerBgColor">表头背景色 Header background color</param>
    /// <param name="headerTextColor">表头文本色 Header text color</param>
    /// <param name="lineHeight">行高 Line height</param>
    /// <param name="tableSize">
    /// 表格尺寸，用于设置表格的行高和列宽，列宽分为最小列宽和最大列宽。<br /><br />
    /// Table size, used to set the row height and column width of the table, and the column width is divided into the minimum column width and the maximum column width.
    /// </param>
    public TableStyle(Color headerBgColor, Color headerTextColor, decimal lineHeight, TableSize tableSize)
    {
        HeaderBgColor = headerBgColor;
        HeaderTextColor = headerTextColor;
        LineHeight = lineHeight;
        MinColumnWidth = tableSize.MinColumnWidth;
        MaxColumnWidth = tableSize.MaxColumnWidth;
        FontSize = tableSize.FontSize;
    }

    /// <summary>
    /// 表头背景色 Header background color
    /// </summary>
    public Color HeaderBgColor { get; set; }

    /// <summary>
    /// 表头文本色 Header text color
    /// </summary>
    public Color HeaderTextColor { get; set; }

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

    /// <summary>
    /// 自动筛选 Auto filter
    /// </summary>
    public decimal AutoFilter { get; set; }
}
