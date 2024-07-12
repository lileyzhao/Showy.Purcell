using System;
using System.Collections.Generic;
using System.Text;

namespace Showy.Purcell;

/// <summary>
/// 预设的表格尺寸 Preset table sizes
/// </summary>
public static class TableSizes
{
    /// <summary>
    /// 迷你 Mini
    /// </summary>
    public static TableSize Mini = new TableSize(10, 30, 10, 30);

    /// <summary>
    /// 小 Small
    /// </summary>
    public static TableSize Small = new TableSize(20, 40, 20, 40);

    /// <summary>
    /// 中 Medium
    /// </summary>
    public static TableSize Medium = new TableSize(30, 50, 30, 50);

    /// <summary>
    /// 大 Large
    /// </summary>
    public static TableSize Large = new TableSize(40, 60, 40, 60);

    /// <summary>
    /// 巨大 Huge
    /// </summary>
    public static TableSize Huge = new TableSize(50, 70, 50, 70);
}