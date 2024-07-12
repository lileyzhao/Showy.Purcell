using System;
using System.Text.RegularExpressions;

namespace Showy.Purcell;

/// <summary>
/// 表头配置
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelHeader : Attribute
{
    private int _index = -1;

    /// <inheritdoc cref="ExcelHeader"/>
    /// <param name="index">列索引，从0开始（即列A=0）</param>
    public ExcelHeader(int index)
    {
        if (index is < 0 or > PurcellConst.XlsxMaxColumn - 1)
            throw new ArgumentException($"列索引超过范围 [ 0 - {PurcellConst.XlsxMaxColumn - 1} ZYX ]", nameof(Index));
        _index = index;
    }

    /// <inheritdoc cref="ExcelHeader"/>
    /// <param name="columnNames">列名，支持配置多个，<br />支持使用$A-$ZYX设置为字母索引，但只有一个值且字母均为大写的时候的时候生效。</param>
    public ExcelHeader(params string[] columnNames)
    {
        if (columnNames.Length == 1 && Regex.IsMatch(columnNames[0], @"^\$([A-Z]{1,3})$"))
            _index = ExcelUtility.ColumnLetterToIndex(Regex.Match(columnNames[0], @"^\$([A-Z]{1,3})$").Groups[1].Value);
        else
            ColumnNames = columnNames;
    }

    /// <summary>
    /// 属性名<br /><br />
    /// ！大小写不敏感<br />
    /// ！建议使用 <c>nameof(Student.RealName)</c> 获取<br />
    /// </summary>
    public string? PropertyName { set; get; }

    /// <summary>
    /// 表头列索引，从0开始（即列A=0）。<br /><br />
    /// ！默认值为 -1 表示未设置列索引；
    /// ！索引范围 Xls=[ 0 - 255 ]  Xlsx=[ 0 - 16383 ]；<br />
    /// ！列的定位设置优先级：<see cref="Index"/> > <see cref="IndexName"/> > <see cref="ColumnNames"/>。
    /// </summary>
    public int Index
    {
        get => _index;
        set
        {
            if (value is < -1 or > PurcellConst.XlsxMaxColumn - 1)
                throw new ArgumentException($"列索引超过范围 [ -1 - {PurcellConst.XlsxMaxColumn - 1} ZYX ](其中 -1 表示未设置)", nameof(Index));
            _index = value;
        }
    }

    /// <summary>
    /// 表头列索引名称，例如 "A" 表示索引为 0 的列。<br /><br />
    /// ！列的定位设置优先级：<see cref="Index"/> > <see cref="IndexName"/> > <see cref="ColumnNames"/>。
    /// </summary>
    public string? IndexName
    {
        get => _index == -1 ? null : ExcelUtility.ColumnIndexToLetter(Index);
        set => _index = string.IsNullOrWhiteSpace(value) ? -1 : ExcelUtility.ColumnLetterToIndex(value);
    }

    /// <summary>
    /// 多个表头列名，用于需要适配不同列名的情况，例如 ["姓名", "名称", "名字"]。<br /><br />
    /// ！包含英文字母时大小写不敏感；<br />
    /// ！列的定位设置优先级：<see cref="Index"/> > <see cref="IndexName"/> > <see cref="ColumnNames"/>。
    /// </summary>
    public string[]? ColumnNames { get; set; }

    /// <summary>
    /// 是否忽略导入，默认值为 false。
    /// </summary>
    public bool IgnoreImport { get; set; }

    /// <summary>
    /// 是否忽略导出，默认值为 false。
    /// </summary>
    public bool IgnoreExport { get; set; }

    /// <summary>
    /// 是否隐藏列，默认值为 false。
    /// </summary>
    public bool ColumnHidden { get; set; }

    /// <summary>
    /// 列宽（字符数，中文或全角等于2个字符宽度）
    /// </summary>
    public double ColumnWidth { get; set; }
}