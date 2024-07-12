namespace Showy.Purcell;

/// <summary>
/// 清除空白字符的模式。(空白字符包含空格和零宽字符)
/// </summary>
public enum ClearWhiteMode
{
    /// <summary>
    /// 不进行ClearWhite操作
    /// </summary>
    None = 0,

    /// <summary>
    /// 对表头进行ClearWhite
    /// </summary>
    ClearHeader = 1,

    /// <summary>
    /// 对单元格值进行ClearWhite
    /// </summary>
    ClearValue = 2,

    /// <summary>
    /// 同时对表头和单元格值进行ClearWhite
    /// </summary>
    ClearAll = 3
}