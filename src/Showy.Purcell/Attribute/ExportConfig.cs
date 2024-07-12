using System;
using System.Collections.Generic;
using System.Text;

namespace Showy.Purcell;

/// <summary>
/// Excel导出配置<br /><br />
/// Excel export configuration
/// </summary>
public class ExportConfig
{
    /*
     * TableSize // 表格大小，可选项：mini、small、medium、large、huge，默认 medium<br />
     * WriteStart，类型 (int startRow, int startColumn) // 写入起始位置，可选项：A1、B2、C3、D4、E5，默认 A1<br />
     * 
     * Purcell.Export(filePath,TableStyles default,TableSize medium,Position writeStart,
     *  bool autoFilter,) // 导出Excel文件，使用默认表格样式<br />
     * 
     * Purcell.Export(TableStyles default,TableSize medium,Position writeStart,
     *  bool autoFilter) // 导出Excel文件不落地返回Stream，使用默认表格样式<br />
     * 
     * 
     * 
     */


    /// <summary>
    /// 表格样式
    /// </summary>
    public TableStyle TableStyle { get; set; } = TableStyles.Default;
}