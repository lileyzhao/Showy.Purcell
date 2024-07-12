using System.Reflection;
using System.Text.RegularExpressions;
using Showy.Purcell.Tests.Dto;

namespace Showy.Purcell.Tests;

public class QueryGenericFileStreamDefault : QueryBase
{
    [Fact]
    public void QueryGenericHasHeader_OR1()
    {
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryGeneric)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var excelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<EmployeeHasHeader>();
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query<EmployeeHasHeader>(stream, excelFileType))
                {
                    Assert.True(item.EmpId > 0);
                    Assert.True(Convert.ToString(item.RealName ?? string.Empty).StartsWith(Convert.ToString(item.LastName ?? string.Empty)));
                    list.Add(item);
                }
            }

            Assert.Equal(10, list.Count);
        }
    }

    [Fact]
    public void QueryGenericHasHeader_OR2()
    {
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryGeneric)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var excelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<EmployeeHasHeader>();
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query<EmployeeHasHeader>(stream, excelFileType, 0))
                {
                    Assert.True(item.EmpId > 0);
                    Assert.True(Convert.ToString(item.RealName ?? string.Empty).StartsWith(Convert.ToString(item.LastName ?? string.Empty)));
                    list.Add(item);
                }
            }

            Assert.Equal(10, list.Count);
        }
    }

    [Fact]
    public void QueryGenericHasHeader_OR3()
    {
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryGeneric)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var excelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<EmployeeHasHeader>();
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query<EmployeeHasHeader>(stream, excelFileType, SheetIndex.Sheet1))
                {
                    Assert.True(item.EmpId > 0);
                    Assert.True(Convert.ToString(item.RealName ?? string.Empty).StartsWith(Convert.ToString(item.LastName ?? string.Empty)));
                    list.Add(item);
                }
            }

            Assert.Equal(10, list.Count);
        }
    }

    [Fact]
    public void QueryGenericHasHeader_OR4()
    {
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryGeneric)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var excelFileType = file.EndsWith("csv") ? ExcelFileType.Csv : file.EndsWith("xls") ? ExcelFileType.Xls : ExcelFileType.Xlsx;

            var list = new List<EmployeeHasHeader>();
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (var item in Purcell.Query<EmployeeHasHeader>(stream, excelFileType, "Sheet1"))
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