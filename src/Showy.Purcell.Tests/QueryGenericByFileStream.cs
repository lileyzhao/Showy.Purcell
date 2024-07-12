using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Showy.Purcell.Tests.Dto;

namespace Showy.Purcell.Tests;

public class QueryGenericByFileStream : QueryBase
{
    [Fact]
    public void QueryGenericNoHeader_Care1()
    {
        var excelSetting = new ExcelSetting { SheetIndex = (int)SheetIndex.Sheet1, HeaderRow = -1, DataStartRow = 0, IgnoreDataError = false };
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryGeneric)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var encoding = file.EndsWith("csv") ? Encoding.GetEncoding("GB2312") : Encoding.Default;
            excelSetting.Encoding = encoding;
            excelSetting.ExcelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<EmployeeNoHeader1>();

            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query<EmployeeNoHeader1>(stream, excelSetting))
                {
                    Assert.True(item.EmpId > 0);
                    Assert.True(Convert.ToString(item.XingMing ?? string.Empty).StartsWith(Convert.ToString(item.Xing ?? string.Empty)));
                    list.Add(item);
                }
            }

            Assert.Equal(10, list.Count);
        }
    }

    [Fact]
    public void QueryGenericNoHeader_Care2()
    {
        var excelSetting = new ExcelSetting { SheetIndex = (int)SheetIndex.Sheet1, HeaderRow = -1, DataStartRow = 0, IgnoreDataError = false };
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryGeneric)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var encoding = file.EndsWith("csv") ? Encoding.GetEncoding("GB2312") : Encoding.Default;
            excelSetting.Encoding = encoding;
            excelSetting.ExcelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<EmployeeNoHeader2>();

            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query<EmployeeNoHeader2>(stream, excelSetting))
                {
                    Assert.True(item.EmpId > 0);
                    Assert.True(Convert.ToString(item.XingMing ?? string.Empty).StartsWith(Convert.ToString(item.Xing ?? string.Empty)));
                    list.Add(item);
                }
            }

            Assert.Equal(10, list.Count);
        }
    }

    [Fact]
    public void QueryGenericHasHeader_Care1()
    {
        var excelSetting = new ExcelSetting { SheetIndex = (int)SheetIndex.Sheet1, HeaderRow = 0, DataStartRow = 1, IgnoreDataError = false };
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryGeneric)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var encoding = file.EndsWith("csv") ? Encoding.GetEncoding("GB2312") : Encoding.Default;
            excelSetting.Encoding = encoding;
            excelSetting.ExcelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<EmployeeHasHeader>();

            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query<EmployeeHasHeader>(stream, excelSetting))
                {
                    Assert.True(item.EmpId > 0);
                    Assert.True(Convert.ToString(item.RealName ?? string.Empty).StartsWith(Convert.ToString(item.LastName ?? string.Empty)));
                    list.Add(item);
                }
            }

            Assert.Equal(10, list.Count);
        }
    }
}