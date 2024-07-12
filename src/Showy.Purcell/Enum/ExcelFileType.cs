namespace Showy.Purcell;

/// <summary>
/// Excel文件类型
/// </summary>
public enum ExcelFileType
{
    /// <summary>
    /// 未知文件
    /// </summary>
    Unknown = 0,

    /// <summary>
    /// .xls文件
    /// </summary>
    Xls = 1,

    /// <summary>
    /// .xlsx文件
    /// </summary>
    Xlsx = 2,

    /// <summary>
    /// .xlsm文件
    /// </summary>
    Xlsm = 3,

    /// <summary>
    /// .xlsb文件
    /// </summary>
    Xlsb = 4,

    /// <summary>
    /// .csv文件
    /// </summary>
    Csv = 5
}