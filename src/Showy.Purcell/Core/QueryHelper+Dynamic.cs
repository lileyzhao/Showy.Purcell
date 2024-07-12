using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Sylvan.Data.Csv;
using Sylvan.Data.Excel;

namespace Showy.Purcell.Core;

internal static partial class QueryHelper
{
    /// <inheritdoc cref="QueryHelper"/>
    internal static async IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(Stream fileStream, ExcelSetting excelSetting, params ExcelHeader[] excelColumns)
    {
        var reader = Preliminary(fileStream, excelSetting);

        var rowIndex = -1;
        var fieldCount = 0;
        string[]? columnsMap = null;
        // 创建一个缓存字典，用于减少每次创建新字典的开销
        var rowCache = new Dictionary<string, object?>();
        while (await reader.ReadAsync())
        {
            rowIndex++;

            // 确定 FieldCount, 并填充列名映射
            if (fieldCount == 0 || excelSetting.HeaderRow == rowIndex || (excelSetting.HeaderRow == -1 && excelSetting.DataStartRow == rowIndex))
            {
                fieldCount = excelSetting.ExcelFileType == ExcelFileType.Csv ? ((CsvDataReader)reader).RowFieldCount : ((ExcelDataReader)reader).RowFieldCount;
                columnsMap = new string[fieldCount];
                for (var i = 0; i < fieldCount; i++)
                {
                    columnsMap[i] = ExcelUtility.ColumnIndexToLetter(i);
                }

                if (excelSetting.HeaderRow == rowIndex)
                {
                    for (var i = 0; i < fieldCount; i++)
                    {
                        var colName = excelSetting.ClearWhiteMode is ClearWhiteMode.ClearHeader or ClearWhiteMode.ClearAll
                            ? PurcellHelper.ClearWhiteSpace(reader.GetString(i)) ?? string.Empty
                            : excelSetting.TrimMode is TrimMode.TrimHeader or TrimMode.TrimAll
                                ? reader.GetString(i).Trim()
                                : reader.GetString(i);

                        if (!string.IsNullOrWhiteSpace(colName) && !columnsMap.Contains(colName))
                        {
                            columnsMap[i] = colName;
                        }

                        foreach (var colSet in excelColumns)
                        {
                            if (!string.IsNullOrWhiteSpace(colSet.PropertyName) && !columnsMap.Contains(colSet.PropertyName) && colSet.ColumnNames?.Contains(colName) == true)
                            {
                                columnsMap[i] = colSet.PropertyName;
                            }
                        }
                    }

                    var repeatedHeaders = columnsMap.GroupBy(h => h).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
                    if (repeatedHeaders.Any())
                    {
                        throw new DuplicateHeaderException($"表格中存在重复的表头（{string.Join('、', repeatedHeaders)}）");
                    }
                }
            }

            if (rowIndex < excelSetting.DataStartRow) continue;

            if (columnsMap == null || columnsMap.Length == 0) throw new DataException(ErrorConst.E822);

            // 开始读取单元格值
            rowCache.Clear(); // 清空字典缓存
            for (var i = 0; i < columnsMap.Length; i++)
            {
                var val = reader.GetValue(i);
                rowCache[columnsMap[i]] = val is DBNull ? null : val;
            }

            // 返回缓存字典的副本
            yield return new Dictionary<string, object?>(rowCache);
        }
    }

    /// <inheritdoc cref="QueryHelper"/>
    internal static IEnumerable<IDictionary<string, object?>> Query(Stream fileStream, ExcelSetting excelSetting, params ExcelHeader[] excelColumns)
    {
        foreach (var item in QueryAsync(fileStream, excelSetting, excelColumns).ToEnumerable())
            yield return item;
    }
}