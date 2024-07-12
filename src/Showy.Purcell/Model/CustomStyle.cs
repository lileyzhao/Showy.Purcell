using LargeXlsx;
using System;
using System.Collections.Generic;
using System.Text;

namespace Showy.Purcell.Model;

/// <summary>
/// 自定义表格样式 Custom table style
/// </summary>
public class CustomStyle
{
    /// <summary>
    /// 表头样式
    /// </summary>
    public XlsxStyle HeaderStyle { get; set; } = XlsxStyle.Default;

    /// <summary>
    /// 内容样式
    /// </summary>
    public XlsxStyle ContentStyle { get; set; } = XlsxStyle.Default;

    /// <summary>
    /// 表底样式(一般为合计行)
    /// </summary>
    public XlsxStyle FooterStyle { get; set; } = XlsxStyle.Default;
}
