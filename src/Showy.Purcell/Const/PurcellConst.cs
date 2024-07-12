using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Sylvan.Data.Excel;

namespace Showy.Purcell;

/// <summary>
/// 一些常量
/// </summary>
internal static class PurcellConst
{
    /// <summary>
    /// Xlsx文件最大支持的列数
    /// </summary>
    internal const short XlsxMaxColumn = 1 << 14;

    /// <summary>
    /// Xls文件最大支持的列名
    /// </summary>
    internal const short XlsMaxColumn = 1 << 8;

    /// <summary>
    /// Xlsx文件最大支持的Worksheet数量
    /// </summary>
    internal const int XlsxMaxSheet = int.MaxValue;

    /// <summary>
    /// Xls文件最大支持的Worksheet数量
    /// </summary>
    internal const short XlsMaxSheet = 1 << 8;

    /// <summary>
    /// 1899-12-30
    /// </summary>
    internal static readonly DateTime Epoch1900 = new(1899, 12, 30);

    /// <summary>
    /// 1904-01-01
    /// </summary>
    internal static readonly DateTime Epoch1904 = new(1904, 1, 1);

    internal static readonly Regex TimeRegex = new(@"^[0-2]?[0-9]{1}(:[0-5]?[0-9]{1}){2}(.([0-9]?){7})?$");

    internal static readonly Dictionary<ExcelFileType, ExcelWorkbookType> ExcelTypeMapping = new()
    {
        { ExcelFileType.Xls, ExcelWorkbookType.Excel },
        { ExcelFileType.Xlsx, ExcelWorkbookType.ExcelXml },
        { ExcelFileType.Xlsm, ExcelWorkbookType.ExcelXml },
        { ExcelFileType.Xlsb, ExcelWorkbookType.ExcelBinary }
    };

    internal static readonly Dictionary<string, ExcelFileType> ExcelExtMapping = new(StringComparer.OrdinalIgnoreCase)
    {
        { ".xls", ExcelFileType.Xls },
        { ".xlsx", ExcelFileType.Xlsx },
        { ".xlsm", ExcelFileType.Xlsm },
        { ".xlsb", ExcelFileType.Xlsb },
        { ".csv", ExcelFileType.Csv }
    };
}