using SharpCompress.Common;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;

namespace Showy.Purcell;

/**
 * 这里是Purcell查询读取相关的同步方法及其重载方法。<br /><br />
 * There are some sync methods `Purcell.Query` for Purcell query here.
 */
public static partial class Purcell
{
    /// <summary></summary>
    public static IEnumerable<IDictionary<string, object?>> Query(string filePath, int sheetIndex,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null)
    {
        var querySetting = new ExcelSetting(sheetIndex)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        foreach (var item in Query(filePath, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<IDictionary<string, object?>> Query(string filePath, SheetIndex sheetIndex,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null)
    {
        var querySetting = new ExcelSetting(sheetIndex)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        foreach (var item in Query(filePath, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<IDictionary<string, object?>> Query(string filePath, string sheetName,
        int headerRow = 0, int dataStartRow = 1, bool ignoreDataError = true, Encoding? encoding = null)
    {
        var querySetting = new ExcelSetting(sheetName)
        {
            HeaderRow = headerRow,
            DataStartRow = dataStartRow,
            IgnoreDataError = ignoreDataError,
            Encoding = encoding ?? GetEncodingGb2312()
        };

        foreach (var item in Query(filePath, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<IDictionary<string, object?>> Query(Stream fileStream, ExcelFileType excelFileType)
    {
        foreach (var item in Query(fileStream, new ExcelSetting { ExcelFileType = excelFileType }))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<IDictionary<string, object?>> Query(Stream fileStream, ExcelFileType excelFileType, int sheetIndex,
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

        foreach (var item in Query(fileStream, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<IDictionary<string, object?>> Query(Stream fileStream, ExcelFileType excelFileType, SheetIndex sheetIndex,
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

        foreach (var item in Query(fileStream, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<IDictionary<string, object?>> Query(Stream fileStream, ExcelFileType excelFileType, string sheetName,
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

        foreach (var item in Query(fileStream, querySetting))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<T> Query<T>(string filePath, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        foreach (var item in Query<T>(filePath, null, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<T> Query<T>(string filePath, int sheetIndex,
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

        foreach (var item in Query<T>(filePath, querySetting, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<T> Query<T>(string filePath, SheetIndex sheetIndex,
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

        foreach (var item in Query<T>(filePath, querySetting, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<T> Query<T>(string filePath, string sheetName,
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

        foreach (var item in Query<T>(filePath, querySetting, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<T> Query<T>(Stream fileStream, ExcelFileType excelFileType, params ExcelHeader[] excelHeaders)
        where T : class, new()
    {
        var excelSetting = typeof(T).GetCustomAttributes(typeof(ExcelSetting), true).OfType<ExcelSetting>().FirstOrDefault();
        if (excelSetting != null) excelSetting.ExcelFileType = excelFileType;

        foreach (var item in Query<T>(fileStream, excelSetting ?? new ExcelSetting { ExcelFileType = excelFileType }, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<T> Query<T>(Stream fileStream, ExcelFileType excelFileType, int sheetIndex,
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

        foreach (var item in Query<T>(fileStream, querySetting, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<T> Query<T>(Stream fileStream, ExcelFileType excelFileType, SheetIndex sheetIndex,
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

        foreach (var item in Query<T>(fileStream, querySetting, excelHeaders))
            yield return item;
    }

    /// <summary></summary>
    public static IEnumerable<T> Query<T>(Stream fileStream, ExcelFileType excelFileType, string sheetName,
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

        foreach (var item in Query<T>(fileStream, querySetting, excelHeaders))
            yield return item;
    }
}