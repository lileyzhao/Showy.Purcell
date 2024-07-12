namespace Showy.Purcell;

/// <summary>
/// Trim模式
/// </summary>
public enum TrimMode
{
    /// <summary>
    /// 不进行Trim操作
    /// </summary>
    None = 0,

    /// <summary>
    /// 对表头进行Trim
    /// </summary>
    TrimHeader = 1,

    /// <summary>
    /// 对单元格值进行Trim
    /// </summary>
    TrimValue = 2,

    /// <summary>
    /// 同时对表头和单元格值进行Trim
    /// </summary>
    TrimAll = 3
}