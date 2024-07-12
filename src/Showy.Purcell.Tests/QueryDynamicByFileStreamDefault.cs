using System.Reflection;
using System.Text.RegularExpressions;

namespace Showy.Purcell.Tests;

public class QueryDynamicByFileStreamDefault : QueryBase
{
    [Fact]
    public void QueryDynamicHasHeader_OR1()
    {
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryDynamic)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var excelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<IDictionary<string, object?>>();
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query(stream, excelFileType))
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
    public void QueryDynamicHasHeader_OR2()
    {
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryDynamic)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var excelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<IDictionary<string, object?>>();
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query(stream, excelFileType, 0))
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
    public void QueryDynamicHasHeader_OR3()
    {
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryDynamic)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var excelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<IDictionary<string, object?>>();
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query(stream, excelFileType, SheetIndex.Sheet1))
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
    public void QueryDynamicHasHeader_OR4()
    {
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryDynamic)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var excelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<IDictionary<string, object?>>();
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query(stream, excelFileType, "Sheet1"))
                {
                    Assert.True(int.TryParse(item["员工ID"]!.ToString(), out var _));
                    Assert.True(Convert.ToString(item["姓名"])!.StartsWith(Convert.ToString(item["姓氏"])!));
                    list.Add(item);
                }
            }

            Assert.Equal(10, list.Count);
        }
    }
}