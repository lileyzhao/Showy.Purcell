using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Sylvan.Data.Csv;
using Sylvan.Data.Excel;

namespace Showy.Purcell.Core;

/// <summary>
/// Excel读取核心帮助类
/// </summary>
internal static partial class QueryHelper
{
    /// <summary>
    /// 预处理，并返回DbReader对象
    /// </summary>
    /// <param name="fileStream">文件流</param>
    /// <param name="excelSetting">Excel配置</param>
    /// <returns>DbReader对象</returns>
    /// <exception cref="ArgumentException"></exception>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    /// <exception cref="InvalidDataException"></exception>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static DbDataReader Preliminary(Stream fileStream, ExcelSetting excelSetting)
    {
        try
        {
            if (excelSetting.ExcelFileType == ExcelFileType.Csv)
                return CsvDataReader.Create(new StreamReader(fileStream, excelSetting.Encoding), new CsvDataReaderOptions { HasHeaders = false });

            var reader = ExcelDataReader.Create(fileStream, PurcellConst.ExcelTypeMapping[excelSetting.ExcelFileType],
                new ExcelDataReaderOptions { GetErrorAsNull = true, Schema = ExcelSchema.NoHeaders });

            // 移动至目标工作簿。！优先级：SheetName > SheetIndex，SheetName无效时尝试使用SheetIndex。
            var sheetNames = reader.WorksheetNames.ToList();
            var sheetNameIndex = string.IsNullOrEmpty(excelSetting.SheetName) ? -1 : sheetNames.IndexOf(excelSetting.SheetName);
            if (sheetNameIndex != -1)
            {
                // 注意For体 i<excelSetting.SheetIndex 表示当 excelSetting.SheetIndex=0 时不会执行 NextResult() 进行步进
                // 因为Reader.Create时，默认处于第一个WorkSheet位置
                for (var i = 0; i < sheetNameIndex; i++)
                    reader.NextResult(); // 步进到对应的Sheet位置
            }
            else
            {
                if (excelSetting.SheetIndex < 0) throw new ArgumentException(ErrorConst.E812, nameof(ExcelSetting.SheetIndex));
                if (excelSetting.SheetIndex + 1 > sheetNames.Count)
                    throw new ArgumentOutOfRangeException(nameof(ExcelSetting.SheetIndex), ErrorConst.E813);

                // 注意For体 i<excelSetting.SheetIndex 表示当 excelSetting.SheetIndex=0 时不会执行 NextResult() 进行步进
                // 因为Reader.Create时，默认处于第一个WorkSheet位置
                for (var i = 0; i < excelSetting.SheetIndex; i++)
                    reader.NextResult(); // 步进到对应的Sheet位置
            }

            return reader;
        }
        catch (InvalidDataException dataException)
        {
            throw new InvalidDataException(ErrorConst.E802, dataException);
        }
        catch (Exception ex)
        {
            throw new InvalidDataException(ErrorConst.E801, ex);
        }
    }

    /// <summary>
    /// 获取所有Worksheet名称
    /// </summary>
    /// <param name="fileStream">文件流</param>
    /// <param name="excelFileType">Excel文件类型</param>
    /// <returns>Worksheet名称列表</returns>
    internal static IEnumerable<string> GetSheetNames(Stream fileStream, ExcelFileType excelFileType)
    {
        switch (excelFileType)
        {
            case ExcelFileType.Csv:
                yield return "Sheet1";
                break;
            case ExcelFileType.Xls:
            case ExcelFileType.Xlsx:
            case ExcelFileType.Xlsm:
            case ExcelFileType.Xlsb:
                {
                    using var reader = ExcelDataReader.Create(fileStream, PurcellConst.ExcelTypeMapping[excelFileType],
                        new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders, GetErrorAsNull = true });
                    foreach (var name in reader.WorksheetNames) yield return name;
                    break;
                }
            default:
                throw new InvalidDataException(ErrorConst.E803);
        }
    }
}