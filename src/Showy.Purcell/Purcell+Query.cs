using Showy.Purcell.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Showy.Purcell;

/**
 * 这里是Purcell查询读取相关的通用方法或基础方法。<br /><br />
 * There are some common basic methods for Purcell.Query here.
 */
public static partial class Purcell
{
    /// <summary>
    /// <inheritdoc cref="Purcell"/>
    /// 通过文件路径读取Excel文件数据。
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="excelSetting">Excel配置</param>
    /// <returns>行数据迭代器</returns>
    /// <exception cref="InvalidDataException"></exception>
    /// <exception cref="ArgumentException"></exception>
    public static IEnumerable<IDictionary<string, object?>> Query(string filePath, ExcelSetting? excelSetting = null)
    {
        excelSetting = ProcExcelSetting(excelSetting, filePath);

        using var fileStream = OpenFile(filePath, excelSetting);
        foreach (var item in Query(fileStream, excelSetting))
            yield return item;
    }

    /// <summary>
    /// <inheritdoc cref="Purcell"/>
    /// 通过文件流读取Excel文件数据。
    /// </summary>
    /// <param name="fileStream">文件流</param>
    /// <param name="excelSetting">Excel配置</param>
    /// <returns>行数据迭代器</returns>
    /// <exception cref="InvalidDataException"></exception>
    /// <exception cref="ArgumentException"></exception>
    public static IEnumerable<IDictionary<string, object?>> Query(Stream fileStream, ExcelSetting? excelSetting = null)
    {
        excelSetting = ProcExcelSetting(excelSetting);

        foreach (var item in QueryHelper.Query(fileStream, excelSetting))
            yield return item;
    }

    /// <summary>
    /// <inheritdoc cref="Purcell"/>
    /// 通过文件路径读取Excel文件数据，可以自定义表头列配置。
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="excelSetting">Excel配置</param>
    /// <param name="excelHeaders">表头列配置</param>
    /// <typeparam name="T">实体泛型</typeparam>
    /// <returns></returns>
    /// <exception cref="ArgumentException"></exception>
    public static IEnumerable<T> Query<T>(string filePath, ExcelSetting? excelSetting, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var newSetting = ProcExcelSetting<T>(excelSetting, filePath);

        using var fileStream = OpenFile(filePath, newSetting);
        foreach (var item in Query<T>(fileStream, newSetting, excelHeaders))
            yield return item;
    }

    /// <summary>
    /// <inheritdoc cref="Purcell"/>
    /// 通过文件流读取Excel文件数据，可以自定义表头列配置。
    /// </summary>
    /// <param name="fileStream">文件流</param>
    /// <param name="excelSetting">Excel配置</param>
    /// <param name="excelHeaders">表头列配置</param>
    /// <typeparam name="T">实体泛型</typeparam>
    /// <returns>行的实例数据迭代器</returns>
    /// <exception cref="ArgumentException"></exception>
    public static IEnumerable<T> Query<T>(Stream fileStream, ExcelSetting? excelSetting, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var newSetting = ProcExcelSetting<T>(excelSetting);

        foreach (var item in QueryHelper.Query<T>(fileStream, newSetting, excelHeaders))
            yield return item;
    }

    /// <summary>
    /// <inheritdoc cref="Purcell"/>
    /// 通过文件路径读取Excel文件数据。
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="excelSetting">Excel配置</param>
    /// <returns>行数据迭代器</returns>
    /// <exception cref="InvalidDataException"></exception>
    /// <exception cref="ArgumentException"></exception>
    public static async IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(string filePath, ExcelSetting? excelSetting = null)
    {
        excelSetting = ProcExcelSetting(excelSetting, filePath);

        using var fileStream = OpenFile(filePath, excelSetting);
        await foreach (var item in QueryAsync(fileStream, excelSetting))
            yield return item;
    }

    /// <summary>
    /// <inheritdoc cref="Purcell"/>
    /// 通过文件流读取Excel文件数据。
    /// </summary>
    /// <param name="fileStream">文件流</param>
    /// <param name="excelSetting">Excel配置</param>
    /// <returns>行数据迭代器</returns>
    /// <exception cref="InvalidDataException"></exception>
    /// <exception cref="ArgumentException"></exception>
    public static async IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(Stream fileStream, ExcelSetting? excelSetting = null)
    {
        excelSetting = ProcExcelSetting(excelSetting);

        await foreach (var item in QueryHelper.QueryAsync(fileStream, excelSetting))
            yield return item;
    }

    /// <summary>
    /// <inheritdoc cref="Purcell"/>
    /// 通过文件路径读取Excel文件数据，可以自定义表头列配置。
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <param name="excelSetting">Excel配置</param>
    /// <param name="excelHeaders">表头列配置</param>
    /// <typeparam name="T">实体泛型</typeparam>
    /// <returns></returns>
    /// <exception cref="ArgumentException"></exception>
    public static async IAsyncEnumerable<T> QueryAsync<T>(string filePath, ExcelSetting? excelSetting, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var newSetting = ProcExcelSetting<T>(excelSetting, filePath);

        using var fileStream = OpenFile(filePath, newSetting);
        await foreach (var item in QueryAsync<T>(fileStream, newSetting, excelHeaders))
            yield return item;
    }

    /// <summary>
    /// <inheritdoc cref="Purcell"/>
    /// 通过文件流读取Excel文件数据，可以自定义表头列配置。
    /// </summary>
    /// <param name="fileStream">文件流</param>
    /// <param name="excelSetting">Excel配置</param>
    /// <param name="excelHeaders">表头列配置</param>
    /// <typeparam name="T">实体泛型</typeparam>
    /// <returns>行的实例数据迭代器</returns>
    /// <exception cref="ArgumentException"></exception>
    public static async IAsyncEnumerable<T> QueryAsync<T>(Stream fileStream, ExcelSetting? excelSetting, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var newSetting = ProcExcelSetting<T>(excelSetting);

        await foreach (var item in QueryHelper.QueryAsync<T>(fileStream, newSetting, excelHeaders))
            yield return item;
    }

    /// <summary>
    /// 获取所有Worksheet名称
    /// </summary>
    /// <param name="filePath">文件路径</param>
    /// <returns>Worksheet名称列表</returns>
    public static List<string> GetSheetNames(string filePath)
    {
        var excelSetting = ProcExcelSetting(null, filePath);

        using var fileStream = OpenFile(filePath, excelSetting);
        return GetSheetNames(fileStream, excelSetting.ExcelFileType);
    }

    /// <summary>
    /// 获取所有Worksheet名称
    /// </summary>
    /// <param name="fileStream">文件流</param>
    /// <param name="excelFileType">Excel文件类型</param>
    /// <returns>Worksheet名称列表</returns>
    public static List<string> GetSheetNames(Stream fileStream, ExcelFileType excelFileType)
    {
        return QueryHelper.GetSheetNames(fileStream, excelFileType).ToList();
    }

    /// <summary>
    /// 获取Encoding.GB2312
    /// </summary>
    /// <returns>Encode.GB2312</returns>
    public static Encoding GetEncodingGb2312()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding("GB2312");
    }

    /// <summary>
    /// 检查ExcelSetting
    /// </summary>
    /// <param name="excelSetting"></param>
    /// <param name="filePath"></param>
    /// <returns></returns>
    /// <exception cref="InvalidDataException"></exception>
    /// <exception cref="ArgumentException"></exception>
    private static ExcelSetting ProcExcelSetting(ExcelSetting? excelSetting, string? filePath = null)
    {
        excelSetting ??= new ExcelSetting();

        // 设置Excel文件类型
        if (!string.IsNullOrEmpty(filePath) && excelSetting.ExcelFileType == ExcelFileType.Unknown)
            excelSetting.ExcelFileType = PurcellConst.ExcelExtMapping.GetValueOrDefault(Path.GetExtension(filePath).ToLower(), ExcelFileType.Unknown);

        if (excelSetting is { ExcelFileType: ExcelFileType.Unknown })
            throw new InvalidDataException(ErrorConst.E803);

        if (excelSetting.DataStartRow <= excelSetting.HeaderRow)
            throw new ArgumentException(ErrorConst.E821, nameof(ExcelSetting.DataStartRow));

        return excelSetting;
    }

    /// <summary>
    /// 检查ExcelSetting
    /// </summary>
    /// <param name="excelSetting"></param>
    /// <param name="filePath"></param>
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    private static ExcelSetting ProcExcelSetting<T>(ExcelSetting? excelSetting, string? filePath = null)
    {
        excelSetting ??= typeof(T).GetCustomAttributes(typeof(ExcelSetting), true).OfType<ExcelSetting>().FirstOrDefault();

        return ProcExcelSetting(excelSetting, filePath);
    }

    /// <summary>
    /// 读取文件返回Stream
    /// </summary>
    /// <param name="filePath"></param>
    /// <param name="excelSetting"></param>
    /// <returns></returns>
    private static Stream OpenFile(string filePath, ExcelSetting excelSetting)
    {
        return excelSetting.ExcelFileType == ExcelFileType.Csv
            ? new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1024 * 10, FileOptions.SequentialScan)
            : new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1024 * 4);
    }
}