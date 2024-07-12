using System.Reflection;
using System.Text.RegularExpressions;
using Showy.Purcell.Tests.Dto;

namespace Showy.Purcell.Tests;

public class QueryGenericByFilePath : QueryBase
{
    [Fact]
    public void QueryGenericNoHeader_Care1()
    {
        var excelSetting = new ExcelSetting { SheetIndex = (int)SheetIndex.Sheet1, HeaderRow = -1, DataStartRow = 0, IgnoreDataError = false };
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryGeneric)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            excelSetting.ExcelFileType = ExcelFileType.Unknown; // 设置为未知，自动由Purcell获取文件类型

            var list = new List<EmployeeNoHeader1>();
            foreach (var item in Purcell.Query<EmployeeNoHeader1>(file, excelSetting))
            {
                Assert.True(item.EmpId > 0);
                Assert.True(Convert.ToString(item.XingMing ?? string.Empty).StartsWith(Convert.ToString(item.Xing ?? string.Empty)));
                list.Add(item);
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
            excelSetting.ExcelFileType = ExcelFileType.Unknown; // 设置为未知，自动由Purcell获取文件类型

            var list = new List<EmployeeNoHeader2>();
            foreach (var item in Purcell.Query<EmployeeNoHeader2>(file, excelSetting))
            {
                Assert.True(item.EmpId > 0);
                Assert.True(Convert.ToString(item.XingMing ?? string.Empty).StartsWith(Convert.ToString(item.Xing ?? string.Empty)));
                list.Add(item);
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
            excelSetting.ExcelFileType = ExcelFileType.Unknown; // 设置为未知，由Purcell自动获取文件类型

            var list = new List<EmployeeHasHeader>();
            foreach (var item in Purcell.Query<EmployeeHasHeader>(file, excelSetting))
            {
                Assert.True(item.EmpId > 0);
                Assert.True(Convert.ToString(item.RealName ?? string.Empty).StartsWith(Convert.ToString(item.LastName ?? string.Empty)));
                list.Add(item);
            }

            Assert.Equal(10, list.Count);

            // 输出到d盘
            var newList = new List<EmployeeHasHeader>();
            for(int i = 0; i < 10000; i++)
            {
                newList.AddRange(list);
            }
            Purcell.Export(@"D:\123.xlsx", newList);
        }
    }
}