using SharpCompress;
using System;
using System.Collections.Generic;
using System.Text;

namespace Showy.Purcell;

/// <summary>
/// Excel查询读取配置，优先级 参数传递 > 类特性 > 默认配置。<br /><br />
/// Excel setting, priority: parameter > attribute > default configuration.
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class ExcelSetting : Attribute
{
    private int _dataStartRow;
    private int _headerRow;
    private int _sheetIndex;
    private (HashSet<string> False, HashSet<string> True) _boolStrings;

    /// <inheritdoc cref="ExcelSetting"/>
    public ExcelSetting() : this(0, null)
    {
    }

    /// <inheritdoc cref="ExcelSetting"/>
    /// <param name="sheetName">
    ///     工作簿名称，优先级：SheetName > SheetIndex，SheetName无效时尝试使用SheetIndex。<br />
    ///     Worksheet name, priority: SheetName > SheetIndex, try to use SheetIndex when SheetName is invalid.
    /// </param>
    public ExcelSetting(string sheetName) : this(0, sheetName)
    {
    }

    /// <inheritdoc cref="ExcelSetting"/>
    /// <param name="sheetIndex">
    ///     工作簿索引枚举，最大值 Sheet256，默认读取 Sheet1，优先级：SheetName > SheetIndex，SheetName无效时尝试使用SheetIndex。<br />
    ///     Worksheet index enumeration, maximum value Sheet256, default read Sheet1, priority: SheetName > SheetIndex, try to use SheetIndex when SheetName is invalid.
    /// </param>
    public ExcelSetting(SheetIndex sheetIndex) : this((int)sheetIndex, null)
    {
    }

    /// <inheritdoc cref="ExcelSetting"/>
    /// <param name="sheetIndex">
    ///     工作簿索引，范围 0 - 255，优先级：SheetName > SheetIndex，SheetName无效时尝试使用SheetIndex。<br />
    ///     Worksheet index, range 0 - 255, priority: SheetName > SheetIndex, try to use SheetIndex when SheetName is invalid.
    /// </param>
    public ExcelSetting(int sheetIndex) : this(sheetIndex, null)
    {
    }

    private ExcelSetting(int sheetIndex, string? sheetName)
    {
        SheetName = sheetName;
        SheetIndex = sheetIndex;
        HeaderRow = 0;
        DataStartRow = 1;
        ExcelFileType = ExcelFileType.Unknown;
        TrimMode = TrimMode.TrimHeader;
        ClearWhiteMode = ClearWhiteMode.ClearHeader;
        BoolStrings = (
           new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "0", "false", "no", "not", "failure", "否", "错误", "没有" },
           new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "1", "true", "yes", "ok", "success", "是", "正确", "有" }
        );
    }

    /// <summary>
    /// 工作簿索引，从0开始，默认值为0<br /><br />
    /// Worksheet index, starting from 0, default value is 0.
    /// </summary>
    public int SheetIndex
    {
        get => _sheetIndex;
        set
        {
            if (value is < 0 or > 255) throw new ArgumentOutOfRangeException(nameof(SheetIndex), ErrorConst.E814);
            _sheetIndex = value;
        }
    }

    /// <summary>
    /// 工作簿名称。优先级：SheetName > SheetIndex，SheetName无效时尝试使用SheetIndex。<br /><br />
    /// Worksheet name. Priority: SheetName > SheetIndex, try to use SheetIndex when SheetName is invalid.
    /// </summary>
    public string? SheetName { get; set; }

    /// <summary>
    /// 表头行索引，从0开始，默认值为 0，值为 -1 时表示没有表头。<br /><br />
    /// Header row index, starting from 0, default value is 0, -1 means no header.
    /// </summary>
    public int HeaderRow
    {
        get => _headerRow;
        set
        {
            if (value < -1) throw new ArgumentOutOfRangeException(nameof(HeaderRow), $"{nameof(HeaderRow)} 必须 >= -1");
            _headerRow = value;
        }
    }

    /// <summary>
    /// 数据起始行索引，从0开始，默认值为 1。<br /><br />
    /// Data start row index, starting from 0, default value is 1.
    /// </summary>
    public int DataStartRow
    {
        get => _dataStartRow;
        set
        {
            if (value < 0)
                throw new ArgumentOutOfRangeException(nameof(DataStartRow), $"{nameof(DataStartRow)} 必须 >= 0");
            _dataStartRow = value;
        }
    }

    /// <summary>
    /// 忽略单元格的数据错误(不抛出异常)，仅使用Query&lt;T&gt;泛型查询的时候有效。<br /><br />
    /// Ignore data errors of cells (do not throw exceptions), only valid when using Query&lt;T&gt; generic query.
    /// </summary>
    public bool IgnoreDataError { get; set; } = true;

    /// <summary>
    /// 编码格式，仅读取Csv时使用，默认为 "GB2312"。<br /><br />
    /// Encoding format, only used when reading Csv, default is "GB2312".
    /// </summary>
    public Encoding Encoding { get; set; } = Purcell.GetEncodingGb2312();

    /// <summary>
    /// Excel文件类型，读取Stream文件流时需设置，使用文件路径时无需配置。<br /><br />
    /// Excel file type, need to be set when reading Stream file stream, no need to configure when using file path.
    public ExcelFileType ExcelFileType { get; set; }

    /// <summary>
    /// Trim模式过滤文本首尾的空格或空白字符，默认值为 TrimMode.TrimHeader。<br /><br />
    /// Trim mode filters leading and trailing spaces or white characters of text, default value is TrimMode.TrimHeader.
    /// </summary>
    public TrimMode TrimMode { get; set; }

    /// <summary>
    /// 清除空白字符的模式(空白字符包含空格和零宽字符)，默认值为 ClearWhiteMode.ClearHeader。<br /><br />
    /// Clear white character mode (white character includes space and zero-width character), default value is ClearWhiteMode.ClearHeader.
    /// </summary>
    public ClearWhiteMode ClearWhiteMode { get; set; }

    /// <summary>
    /// 布尔类型自定义字符串集合，第一个集合为 False，第二个集合为 True。<br /><br />
    /// Boolean type custom string array, the first array is False, the second array is True.<br /><br />
    /// Defaults: <br />(["0", "false", "no", "not", "failure", "否", "错误", "没有"], ["1", "true", "yes", "ok", "success", "是", "正确", "有"])
    /// </summary>
    public (HashSet<string> False, HashSet<string> True) BoolStrings
    {
        get => _boolStrings;
        set
        {
            // BoolStrings.True
            if (_boolStrings.True == null)
                _boolStrings.True = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            else
                _boolStrings.True.Clear();
            // BoolStrings.False
            if (_boolStrings.False == null)
                _boolStrings.False = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            else
                _boolStrings.False.Clear();
            // 保证大小写不敏感
            value.False.ForEach(s => _boolStrings.False.Add(s));
            value.True.ForEach(s => _boolStrings.True.Add(s));
        }
    }


    /// <summary>
    /// 枚举字典，Key为枚举类型，Value为匹配枚举的字符串列表。<br /><br />
    /// Enum dictionary, Key is the enumeration type, Value is the list of strings that match the enumeration.
    /// </summary>
    public Dictionary<Enum, List<string>>? EnumDictionary { get; set; }
}