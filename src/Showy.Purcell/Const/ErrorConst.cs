namespace Showy.Purcell;

/// <summary>
/// 错误信息常量
/// </summary>
internal static class ErrorConst
{
    /// <summary>
    /// 文件读取失败，请确认文件格式与扩展名一致
    /// </summary>
    internal const string E801 = "文件读取失败，请确认文件格式与扩展名一致（详细错误请参考InnerException）";

    /// <summary>
    /// 文件格式错误或者Excel版本过老，无法读取
    /// </summary>
    internal const string E802 = "文件格式错误或者Excel版本过老，无法读取（详细错误请参考InnerException）";

    /// <summary>
    /// 无效的文件类型，请设置正确的 ExcelFileType
    /// </summary>
    internal const string E803 = "无效的文件类型，请设置正确的 ExcelFileType";

    internal const string E811 = "工作簿“{0}”不存在";
    internal const string E812 = $"{nameof(ExcelSetting.SheetIndex)}不能小于0";
    internal const string E813 = $"{nameof(ExcelSetting.SheetIndex)}超过当前文件工作簿数量";
    internal const string E814 = $"工作簿索引超过范围 [ 0 - 255 ]";


    /// <summary>
    /// 数据起始行号 必须大于 表头行号
    /// </summary>
    internal const string E821 = "数据起始行号 必须大于 表头行号";

    internal const string E822 = "没有找到表头行 或 表头行是空白的";


    /// <summary>
    /// 无头表格需配置 ExcelHeader.Index列索引
    /// </summary>
    internal const string E831 = """
                                 如果配置 ExcelSetting.HeaderRow < 0（无列头），则必须使用 ExcelHeader.Index配置列索引！
                                 ( If ExcelSetting.HeaderRow < 0, You must use ExcelHeader.Index to configure the column index!)
                                 """;
}