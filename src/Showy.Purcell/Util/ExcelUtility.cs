using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Showy.Purcell;

/// <summary>
/// Excel工具类 Excel utility class
/// </summary>
internal static class ExcelUtility
{
    public static Dictionary<int, string> GenerateIndexLetter(int maxColumn)
    {
        Dictionary<int, string> indexLetter = new();
        for (var i = 0; i < maxColumn; i++) indexLetter[i] = ColumnIndexToLetter(i);

        return indexLetter;
    }

    /// <summary>
    /// Excel列索引转换为字母形式 Column index to letter
    /// </summary>
    /// <param name="columnIndex">列索引 Column index</param>
    public static string ColumnIndexToLetter(int columnIndex)
    {
        if (columnIndex < 0)
        {
            throw new ArgumentException("列号必须大于或等于零。Column number must be greater than or equal to zero.");
        }

        StringBuilder columnNameBuilder = new StringBuilder();

        while (columnIndex >= 0)
        {
            int remainder = columnIndex % 26;
            columnNameBuilder.Insert(0, Convert.ToChar('A' + remainder));
            columnIndex = columnIndex / 26 - 1;

            if (columnIndex < 0)
            {
                break;
            }
        }

        return columnNameBuilder.ToString();
    }

    /// <summary>
    /// Excel列字母转换为索引形式 Column letter to index
    /// </summary>
    /// <param name="columnLetter">列字母 Column letter</param>
    public static int ColumnLetterToIndex(string columnLetter)
    {
        if (string.IsNullOrEmpty(columnLetter) || !Regex.IsMatch(columnLetter, "^[A-Z]+$"))
        {
            throw new ArgumentException("无效的 Excel 列名。Invalid Excel column name.");
        }

        int columnIndex = 0;
        int power = 1;

        for (int i = columnLetter.Length - 1; i >= 0; i--)
        {
            char c = columnLetter[i];
            columnIndex += (c - 'A' + 1) * power;
            power *= 26;
        }

        return columnIndex - 1;
    }

    /// <summary>
    /// 解析单元格引用，获取行索引和列索引，比如B3的行索引是2，列索引是 1<br /><br />
    /// Parse the cell reference, get the row index and column index, such as B3, the row index is 2, and the column index is 1
    /// </summary>
    /// <param name="cellReference">单元格引用 Cell reference</param>
    /// <param name="rowIndex">行索引 Row index</param>
    /// <param name="columnIndex">列索引 Column index</param>
    public static bool ParseCellReference(string cellReference, out int rowIndex, out int columnIndex)
    {
        if (string.IsNullOrEmpty(cellReference) || !Regex.IsMatch(cellReference, @"^[A-Z]+[0-9]+$"))
        {
            rowIndex = 0;
            columnIndex = 0;
            return false;
        }

        int index = 0;
        while (char.IsLetter(cellReference[index]))
        {
            index++;
        }

        string columnPart = cellReference.Substring(0, index);
        string rowPart = cellReference.Substring(index);

        columnIndex = ColumnLetterToIndex(columnPart);
        rowIndex = int.Parse(rowPart) - 1;

        return true;
    }
}