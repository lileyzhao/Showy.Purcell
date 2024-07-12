using SharpCompress.Common;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;

namespace Showy.Purcell;

/**
 * 这里是Purcell查询读取相关的异步方法及其重载方法。<br /><br />
 * There are some async methods `Purcell.QueryAsync` for Purcell query here.
 */
public static partial class Purcell
{
    /// <summary></summary>
    public static async IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(string filePath, int sheetIndex,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null)
    {
        var querySetting = new ExcelSetting(sheetIndex)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync(filePath, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(string filePath, SheetIndex sheetIndex,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null)
    {
        var querySetting = new ExcelSetting(sheetIndex)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync(filePath, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(string filePath, string sheetName,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null)
    {
        var querySetting = new ExcelSetting(sheetName)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync(filePath, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(Stream fileStream, ExcelFileType excelFileType)
    {
        await foreach (var item in QueryAsync(fileStream, new ExcelSetting { ExcelFileType = excelFileType }))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(Stream fileStream, ExcelFileType excelFileType, int sheetIndex,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null)
    {
        var querySetting = new ExcelSetting(sheetIndex)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            ExcelFileType = excelFileType,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync(fileStream, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(Stream fileStream, ExcelFileType excelFileType, SheetIndex sheetIndex,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null)
    {
        var querySetting = new ExcelSetting(sheetIndex)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            ExcelFileType = excelFileType,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync(fileStream, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(Stream fileStream, ExcelFileType excelFileType, string sheetName,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null)
    {
        var querySetting = new ExcelSetting(sheetName)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            ExcelFileType = excelFileType,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync(fileStream, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<T> QueryAsync<T>(string filePath, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        await foreach (var item in QueryAsync<T>(filePath, null, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<T> QueryAsync<T>(string filePath, int sheetIndex,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var querySetting = new ExcelSetting(sheetIndex)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync<T>(filePath, querySetting, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<T> QueryAsync<T>(string filePath, SheetIndex sheetIndex,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var querySetting = new ExcelSetting(sheetIndex)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync<T>(filePath, querySetting, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<T> QueryAsync<T>(string filePath, string sheetName,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var querySetting = new ExcelSetting(sheetName)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync<T>(filePath, querySetting, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<T> QueryAsync<T>(Stream fileStream, ExcelFileType excelFileType, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var excelSetting = typeof(T).GetCustomAttributes(typeof(ExcelSetting), true).OfType<ExcelSetting>().FirstOrDefault();
        if (excelSetting != null) excelSetting.ExcelFileType = excelFileType;

        await foreach (var item in QueryAsync<T>(fileStream, excelSetting ?? new ExcelSetting { ExcelFileType = excelFileType }, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<T> QueryAsync<T>(Stream fileStream, ExcelFileType excelFileType, int sheetIndex,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var querySetting = new ExcelSetting(sheetIndex)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            ExcelFileType = excelFileType,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync<T>(fileStream, querySetting, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<T> QueryAsync<T>(Stream fileStream, ExcelFileType excelFileType, SheetIndex sheetIndex,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var querySetting = new ExcelSetting(sheetIndex)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            ExcelFileType = excelFileType,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync<T>(fileStream, querySetting, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static async IAsyncEnumerable<T> QueryAsync<T>(Stream fileStream, ExcelFileType excelFileType, string sheetName,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var querySetting = new ExcelSetting(sheetName)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            ExcelFileType = excelFileType,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        await foreach (var item in QueryAsync<T>(fileStream, querySetting, excelHeaders))
            yield return item;
    }
}