using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Sylvan.Data.Csv;
using Sylvan.Data.Excel;

namespace Showy.Purcell.Core;

internal static partial class QueryHelper
{
    // 缓存类型属性信息
    private static readonly Dictionary<Type, PropertyInfo[]> PropertyCache = new();

    /// <inheritdoc cref="QueryHelper"/>
    internal static async IAsyncEnumerable<T> QueryAsync<T>(Stream fileStream, ExcelSetting excelSetting, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        // dynamic 或 object 类型 不使用当前强类型处理，调用Dictionary返回类型的Query
        if (typeof(T) == typeof(object))
        {
            await foreach (var item in QueryAsync(fileStream, excelSetting))
                // 使用 new ExpandoObject() 创建一个dynamic对象并把item的键值累加赋值给dynamic对象
                yield return (T)item.Aggregate(new ExpandoObject() as IDictionary<string, object?>, (dict, p) =>
                {
                    dict.Add(p);
                    return dict;
                });

            yield break;
        }

        var reader = Preliminary(fileStream, excelSetting);

        // 合并特性和动态配置，动态配置会覆盖特性
        var columnInfos = GetColumnInfos<T>(excelHeaders).ToList();

        if (excelSetting.HeaderRow == -1 && columnInfos.All(h => h.Index < 0))
            throw new ArgumentException(ErrorConst.E831);

        var rowIndex = -1;
        var fieldCount = 0;
        while (await reader.ReadAsync())
        {
            rowIndex++;

            // 映射表头索引
            MapHeadersIndex(reader, excelSetting, columnInfos, ref fieldCount, ref rowIndex);

            if (rowIndex < excelSetting.DataStartRow) continue;

            // 读取数据行
            T dataItem = new();
            for (var colIndex = 0; colIndex < fieldCount; colIndex++)
            {
                var matchedColumnInfos = columnInfos.Where(h => h.Index == colIndex && h.IgnoreImport == false).ToList();
                if (matchedColumnInfos.Count == 0) continue;

                try
                {
                    foreach (var matched in matchedColumnInfos)
                    {

                        SetPropValue(dataItem, matched, reader, excelSetting, ref colIndex);
                    }
                }
                catch (Exception ex) when (ex is InvalidCastException or FormatException or OverflowException or ArgumentNullException)
                {
                    if (!excelSetting.IgnoreDataError)
                        throw new Exception($"cell data error position in {ExcelUtility.ColumnIndexToLetter(colIndex)}{rowIndex}, {ex.Message}", ex);
                }
            }

            yield return dataItem;
        }
    }

    /// <inheritdoc cref="QueryHelper"/>
    internal static IEnumerable<T> Query<T>(Stream fileStream, ExcelSetting excelSetting, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        return QueryAsync<T>(fileStream, excelSetting, excelHeaders).ToEnumerable();
    }

    private static IEnumerable<ExcelColumnInfo> GetColumnInfos<T>(ExcelHeader[] excelColumns)
    {
        if (!PropertyCache.TryGetValue(typeof(T), out var typeProperties))
        {
            typeProperties = typeof(T).GetProperties();
            PropertyCache[typeof(T)] = typeProperties;
        }

        if (typeProperties.Length == 0) yield break; // 没有属性的类直接break

        // 合并特性和配置，配置优先，配置会覆盖特性
        foreach (var propertyInfo in typeProperties)
        {
            // 匹配属性名的时候不区分大小写
            var matchedColumn = excelColumns.FirstOrDefault(q => q.PropertyName?.Equals(propertyInfo.Name, StringComparison.OrdinalIgnoreCase) ?? false);

            if (matchedColumn != null) matchedColumn.PropertyName = propertyInfo.Name; // PropertyName可能大小写不一致，重新赋值准确的属性名
            else
            {
                // 如果不存在特性，则创建一个ExcelColumn对象
                matchedColumn = propertyInfo.GetAttribute<ExcelHeader>() ??
                                new ExcelHeader(propertyInfo.Name, propertyInfo.GetAttribute<DisplayNameAttribute>()?.DisplayName ?? propertyInfo.Name);
                // 如果DisplayName为空字符串，则使用PropertyName
                if (matchedColumn.ColumnNames == null || matchedColumn.ColumnNames.Length == 0) matchedColumn.ColumnNames = [propertyInfo.Name];
            }

            yield return new ExcelColumnInfo(matchedColumn, propertyInfo);
        }
    }

    /// <summary>
    /// 解析头部
    /// </summary>
    /// <param name="reader"></param>
    /// <param name="excelSetting"></param>
    /// <param name="columnInfos"></param>
    /// <param name="fieldCount"></param>
    /// <param name="rowIndex"></param>
    /// <exception cref="ArgumentException"></exception>
    private static void MapHeadersIndex(IDataRecord reader, ExcelSetting excelSetting, List<ExcelColumnInfo> columnInfos, ref int fieldCount, ref int rowIndex)
    {
        if (fieldCount != 0 || (excelSetting.HeaderRow != rowIndex && (excelSetting.HeaderRow != -1 || excelSetting.DataStartRow != rowIndex))) return;

        // 确定 FieldCount
        fieldCount = excelSetting.ExcelFileType == ExcelFileType.Csv ? ((CsvDataReader)reader).RowFieldCount : ((ExcelDataReader)reader).RowFieldCount;

        if (rowIndex != excelSetting.HeaderRow) return;

        var headerDict = new Dictionary<string, byte>(StringComparer.OrdinalIgnoreCase);
        for (byte i = 0; i < fieldCount; i++)
        {
            // 根据清除空白符和Trim模式处理
            var colName = excelSetting.ClearWhiteMode is ClearWhiteMode.ClearHeader or ClearWhiteMode.ClearAll
                ? PurcellHelper.ClearWhiteSpace(reader.GetString(i)) ?? string.Empty
                : excelSetting.TrimMode is TrimMode.TrimHeader or TrimMode.TrimAll
                    ? reader.GetString(i).Trim()
                    : reader.GetString(i);

            headerDict[colName] = i;
        }

        // 没有配置索引的列，通过列名判断进行索引配置
        foreach (var columnInfo in columnInfos.Where(h => h.Index == -1))
        {
            if (columnInfo.ColumnNames == null) continue;
            foreach (var colName in columnInfo.ColumnNames)
            {
                if (headerDict.TryGetValue(colName, out var index))
                {
                    columnInfo.Index = index;
                    break;
                }
            }
        }

        columnInfos = columnInfos.Where(h => h.Index >= 0).ToList();
    }

    /// <summary>
    /// 设置对象的属性值
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="dataItem"></param>
    /// <param name="columnInfo"></param>
    /// <param name="reader"></param>
    /// <param name="excelSetting"></param>
    /// <param name="columnIndex"></param>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void SetPropValue<T>(T dataItem, ExcelColumnInfo columnInfo, IDataRecord reader, ExcelSetting excelSetting, ref int columnIndex)
    {
        var value = reader.GetValue(columnIndex);
        var valueStr = reader.GetString(columnIndex).Trim();

        var propertyInfo = columnInfo.PropertyInfo;

        if (columnInfo.NativeType == typeof(bool))
        {
            if (value is null or DBNull)
            {
                propertyInfo.SetValue(dataItem, false);
            }
            else
            {
                if (excelSetting.BoolStrings.False.Contains(valueStr))
                    propertyInfo.SetValue(dataItem, false);
                else if (excelSetting.BoolStrings.True.Contains(valueStr))
                    propertyInfo.SetValue(dataItem, true);
                else if (bool.TryParse(valueStr, out var parsedValue))
                    propertyInfo.SetValue(dataItem, parsedValue);
                else
                    propertyInfo.SetValue(dataItem, false);
            }
        }
        else if (columnInfo.NativeType == typeof(string))
        {
            valueStr = XmlEncoder.DecodeString(valueStr) ?? string.Empty;

            if (excelSetting.ClearWhiteMode is ClearWhiteMode.ClearValue or ClearWhiteMode.ClearAll)
                valueStr = PurcellHelper.ClearWhiteSpace(valueStr) ?? string.Empty;
            else if (excelSetting.TrimMode is TrimMode.TrimValue or TrimMode.TrimAll)
                valueStr = value.ToString()!.Trim();

            propertyInfo.SetValue(dataItem, valueStr);
        }
        else if (value != null && value is not DBNull && columnInfo.NativeType == typeof(DateTime))
        {
            if (double.TryParse(valueStr ?? string.Empty, out _))
            {
                try
                {
                    propertyInfo.SetValue(dataItem, reader.GetDateTime(columnIndex));
                    return;
                }
                catch
                {
                    // ignored
                }
            }

            if (PurcellConst.TimeRegex.IsMatch(value.ToString()!.Trim())) // 00:00:00.fff 格式先添加年份 0000-01-01
                value = DateTime.MinValue.ToString("yyyy-MM-dd " + value.ToString()!.Trim());

            if (DateTime.TryParseExact(value.ToString(), "M/d/yy h:mm tt", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedValue))
                propertyInfo.SetValue(dataItem, parsedValue.AddYears(-100));
            else if (DateTime.TryParse(value.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedValue1))
                propertyInfo.SetValue(dataItem, parsedValue1);
            else if (DateTime.TryParseExact(value.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedValue2))
                propertyInfo.SetValue(dataItem, parsedValue2);
            else
                propertyInfo.SetValue(dataItem, Convert.ChangeType(value, propertyInfo.PropertyType));
        }
        else if (value != null && value is not DBNull && columnInfo.NativeType.IsEnum)
        {
            var valueEnum = Enum.GetValues(columnInfo.NativeType).Cast<object>()
                .FirstOrDefault(ev => ev.ToString() == valueStr || ((int)ev).ToString() == valueStr);
            if (valueEnum != null) propertyInfo.SetValue(dataItem, valueEnum);
            else if (excelSetting.EnumDictionary != null)
            {
                var matchedEnum = excelSetting.EnumDictionary.Where(q => q.Key.GetType() == columnInfo.NativeType)
                     .Where(q => q.Value.Contains(valueStr))
                     .Select(q => q.Key).FirstOrDefault();
                propertyInfo.SetValue(dataItem, matchedEnum);
            }
        }
        else if (value != null && value is not DBNull && columnInfo.NativeType == typeof(Guid))
        {
            if (Guid.TryParse(value.ToString(), out var outGuid))
                propertyInfo.SetValue(dataItem, outGuid);
        }
        else
        {
            if (columnInfo.NativeType.IsPrimitive || columnInfo.NativeType == typeof(decimal))
                propertyInfo.SetValue(dataItem, Convert.ChangeType(value, columnInfo.NativeType));
            else
                propertyInfo.SetValue(dataItem, Convert.ChangeType(value, propertyInfo.PropertyType));
        }
    }
}