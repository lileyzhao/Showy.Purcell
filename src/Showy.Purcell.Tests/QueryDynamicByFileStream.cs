using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace Showy.Purcell.Tests;

public class QueryDynamicByFileStream : QueryBase
{
    [Fact]
    public void QueryDynamicNoHeader_Dict()
    {
        var excelSetting = new ExcelSetting { SheetIndex = (int)SheetIndex.Sheet1, HeaderRow = -1, DataStartRow = 0, IgnoreDataError = false };
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryDynamic)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var encoding = file.EndsWith("csv") ? Encoding.GetEncoding("GB2312") : Encoding.Default;
            excelSetting.Encoding = encoding;
            excelSetting.ExcelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<IDictionary<string, object?>>();

            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query(stream, excelSetting))
                {
                    Assert.True(int.TryParse(item["A"]!.ToString(), out var _));
                    Assert.True(Convert.ToString(item["C"])!.StartsWith(Convert.ToString(item["B"])!));
                    list.Add(item);
                }
            }

            Assert.Equal(10, list.Count);
        }
    }

    [Fact]
    public void QueryDynamicNoHeader_Dynamic()
    {
        var excelSetting = new ExcelSetting { SheetIndex = (int)SheetIndex.Sheet1, HeaderRow = -1, DataStartRow = 0, IgnoreDataError = false };
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryDynamic)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var encoding = file.EndsWith("csv") ? Encoding.GetEncoding("GB2312") : Encoding.Default;
            excelSetting.Encoding = encoding;
            excelSetting.ExcelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<dynamic>();

            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query<dynamic>(stream, excelSetting))
                {
                    Assert.True(int.TryParse(item.A, out int _));
                    Assert.True(Convert.ToString(item.C).StartsWith(Convert.ToString(item.B)));
                    list.Add(item);
                }
            }

            Assert.Equal(10, list.Count);
        }
    }

    [Fact]
    public void QueryDynamicHasHeader_Dict()
    {
        var excelSetting = new ExcelSetting { SheetIndex = (int)SheetIndex.Sheet1, HeaderRow = 0, DataStartRow = 1, IgnoreDataError = false };
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryDynamic)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var encoding = file.EndsWith("csv") ? Encoding.GetEncoding("GB2312") : Encoding.Default;
            excelSetting.Encoding = encoding;
            excelSetting.ExcelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<IDictionary<string, object?>>();

            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query(stream, excelSetting))
                {
                    Assert.True(int.TryParse(item["员工ID"]!.ToString(), out var _));
                    Assert.True(Convert.ToString(item["姓名"])!.StartsWith(Convert.ToString(item["姓氏"])!));
                    list.Add(item);
                }
            }

            Assert.Equal(10, list.Count);
        }
    }

    [Fact]
    public void QueryDynamicHasHeader_Dynamic()
    {
        var excelSetting = new ExcelSetting { SheetIndex = (int)SheetIndex.Sheet1, HeaderRow = 0, DataStartRow = 1, IgnoreDataError = false };
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryDynamic)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var encoding = file.EndsWith("csv") ? Encoding.GetEncoding("GB2312") : Encoding.Default;
            excelSetting.Encoding = encoding;
            excelSetting.ExcelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<dynamic>();

            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query<dynamic>(stream, excelSetting))
                {
                    Assert.True(int.TryParse(item.员工ID, out int _));
                    Assert.True(Convert.ToString((item as IDictionary<string, object?>)!["姓名"])!.StartsWith(Convert.ToString(item.姓氏)));
                    list.Add(item);
                }
            }

            Assert.Equal(10, list.Count);
        }
    }
}