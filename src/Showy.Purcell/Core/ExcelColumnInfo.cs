using System;
using System.Reflection;

namespace Showy.Purcell;

/// <summary>
/// Excel表头列信息
/// </summary>
internal class ExcelColumnInfo : ExcelHeader
{
    /// <summary>
    /// Excel表头信息
    /// </summary>
    public ExcelColumnInfo(ExcelHeader excelColumn, PropertyInfo propertyInfo) : base(propertyInfo.Name)
    {
        Index = excelColumn.Index;
        IndexName = excelColumn.IndexName;
        ColumnNames = excelColumn.ColumnNames;
        IgnoreImport = excelColumn.IgnoreImport;
        PropertyName = excelColumn.PropertyName;
        
        PropertyInfo = propertyInfo;
        NativeType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;
    }

    /// <summary>
    /// 属性元数据
    /// </summary>
    public PropertyInfo PropertyInfo { get; }

    /// <summary>
    /// 原生类型
    /// </summary>
    public Type NativeType { get; }
}