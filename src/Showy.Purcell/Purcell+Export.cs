using LargeXlsx;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Xml.Linq;

namespace Showy.Purcell;

/**
 * 这里是Purcell导出相关的通用方法或基础方法。<br /><br />
 * There are some common basic methods for Purcell export here.
 */
public static partial class Purcell
{
    // TODO:生成Excel文件


    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="dataItems"></param>
    /// <param name="tableStyle"></param>
    /// <param name="medium"></param>
    /// <param name="writeStart"></param>
    /// <param name="autoFilter"></param>
    /// <returns></returns>
    public static void Export<T>(string filePath, IEnumerable<T> dataItems, TableStyle? tableStyle = null, TableSize? medium = null, (int RowIndex, int ColumnIndex)? writeStart = null, bool autoFilter = false)
        where T : class, new()
    {
        var typeProperties = typeof(T).GetProperties();
        List<ExcelHeader> headers = new List<ExcelHeader>();
        foreach (var property in typeProperties)
        {
            var excelHeader = property.GetAttribute<ExcelHeader>() ??
                new ExcelHeader(property.Name, property.GetAttribute<DisplayNameAttribute>()?.DisplayName ?? property.Name);
            excelHeader.PropertyName = property.Name;
            headers.Add(excelHeader);
        }
        var excelSetting = typeof(T).GetAttribute<ExcelSetting>();

        var headerStyle = new XlsxStyle(
            new XlsxFont("Segoe UI", 9, Color.White, bold: true),
            new XlsxFill(Color.FromArgb(0, 0x45, 0x86)),
            XlsxStyle.Default.Border,
            XlsxStyle.Default.NumberFormat,
            new XlsxAlignment(XlsxAlignment.Horizontal.Left, XlsxAlignment.Vertical.Center));

        // TODO:
        using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);
        {
            var columns = new XlsxColumn[headers.Count];
            for (int i = 0; i < headers.Count; i++)
            {
                columns[i] = XlsxColumn.Formatted(width: GetAdjustedTextLength(headers[i].ColumnNames?[0] ?? headers[i].PropertyName ?? ""));
            }
            xlsxWriter.BeginWorksheet("Sheet1", columns: columns);
            // 输出表头
            xlsxWriter.SetDefaultStyle(headerStyle).BeginRow();
            foreach (var header in headers)
            {
                xlsxWriter.Write(header.ColumnNames?[0] ?? header.PropertyName);
            }
            xlsxWriter.SetDefaultStyle(XlsxStyle.Default);
            // 输出内容
            foreach (var item in dataItems)
            {
                xlsxWriter.BeginRow();
                foreach (var header in headers)
                {
                    var value = item.GetType().GetProperty(header.PropertyName)?.GetValue(item);
                    xlsxWriter.Write(value?.ToString());
                }
            }
        }
        xlsxWriter.SetAutoFilter(1, 1, 1, headers.Count);
    }

    /// <summary>
    /// 计算文本的调整后长度，假设除了英文和常见英文标点外的字符宽度是英文字符的两倍。
    /// </summary>
    /// <param name="text">要计算的文本</param>
    /// <returns>调整后的文本长度</returns>
    private static double GetAdjustedTextLength(string text)
    {
        double length = 0;
        foreach (char c in text)
        {
            // 假设除了英文和常见英文标点外的字符宽度是英文字符的两倍
            if (IsWideCharacter(c))
            {
                length += 2;
            }
            else
            {
                length += 1;
            }
        }
        length += 4;
        if (length < 10) length = 10;
        return length;
    }

    /// <summary>
    /// 判断一个字符是否是宽字符（显示宽度为普通字符的两倍）
    /// </summary>
    /// <param name="ch">要判断的字符</param>
    /// <returns>如果是宽字符，返回true；否则返回false。</returns>
    public static bool IsWideCharacter(char ch)
    {
        int codePoint = ch;

        // 判断字符是否在常见的宽字符范围内
        return (codePoint >= 0x1100 && codePoint <= 0x115F) || // Hangul Jamo
               (codePoint == 0x2329 || codePoint == 0x232A) || // Angle brackets
               (codePoint >= 0x2E80 && codePoint <= 0x2FFB) || // CJK Radicals Supplement and Kangxi Radicals
               (codePoint >= 0x3000 && codePoint <= 0x303E) || // CJK Symbols and Punctuation
               (codePoint >= 0x3040 && codePoint <= 0x309F) || // Hiragana
               (codePoint >= 0x30A0 && codePoint <= 0x30FF) || // Katakana
               (codePoint >= 0x3100 && codePoint <= 0x312F) || // Bopomofo
               (codePoint >= 0x3130 && codePoint <= 0x318F) || // Hangul Compatibility Jamo
               (codePoint >= 0x3190 && codePoint <= 0x31EF) || // Kanbun, Bopomofo Extended
               (codePoint >= 0x3200 && codePoint <= 0x32FF) || // Enclosed CJK Letters and Months
               (codePoint >= 0x3300 && codePoint <= 0x33FF) || // CJK Compatibility
               (codePoint >= 0x3400 && codePoint <= 0x4DBF) || // CJK Unified Ideographs Extension A
               (codePoint >= 0x4E00 && codePoint <= 0x9FFF) || // CJK Unified Ideographs
               (codePoint >= 0xA000 && codePoint <= 0xA4CF) || // Yi Syllables
               (codePoint >= 0xAC00 && codePoint <= 0xD7A3) || // Hangul Syllables
               (codePoint >= 0xF900 && codePoint <= 0xFAFF) || // CJK Compatibility Ideographs
               (codePoint >= 0xFE10 && codePoint <= 0xFE19) || // Vertical forms
               (codePoint >= 0xFE30 && codePoint <= 0xFE6F) || // CJK Compatibility Forms
               (codePoint >= 0xFF00 && codePoint <= 0xFF60) || // Fullwidth ASCII variants
               (codePoint >= 0xFFE0 && codePoint <= 0xFFE6) || // Fullwidth symbol variants
               (codePoint >= 0x1F300 && codePoint <= 0x1F64F) || // Emoticons
               (codePoint >= 0x1F900 && codePoint <= 0x1F9FF) || // Supplemental Symbols and Pictographs
               (codePoint >= 0x20000 && codePoint <= 0x2FFFD) || // CJK Unified Ideographs Extension B-C-D-E-F
               (codePoint >= 0x30000 && codePoint <= 0x3FFFD); // CJK Unified Ideographs Extension G
    }
}