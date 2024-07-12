using System.Reflection;
using System.Text.RegularExpressions;
using Showy.Purcell.Tests.Dto;

namespace Showy.Purcell.Tests;

public class QueryGenericByFilePathDefault : QueryBase
{
    [Fact]
    public void QueryGenericHasHeader_OR1()
    {
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryGeneric)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var list = new List<EmployeeHasHeader>();
            foreach (var item in Purcell.Query<EmployeeHasHeader>(file))
            {
                Assert.True(item.EmpId > 0);
                Assert.True(Convert.ToString(item.RealName ?? string.Empty).StartsWith(Convert.ToString(item.LastName ?? string.Empty)));
                list.Add(item);
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
            var list = new List<EmployeeHasHeader>();
            foreach (var item in Purcell.Query<EmployeeHasHeader>(file, 0))
            {
                Assert.True(item.EmpId > 0);
                Assert.True(Convert.ToString(item.RealName ?? string.Empty).StartsWith(Convert.ToString(item.LastName ?? string.Empty)));
                list.Add(item);
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
            var list = new List<EmployeeHasHeader>();
            foreach (var item in Purcell.Query<EmployeeHasHeader>(file, SheetIndex.Sheet1))
            {
                Assert.True(item.EmpId > 0);
                Assert.True(Convert.ToString(item.RealName ?? string.Empty).StartsWith(Convert.ToString(item.LastName ?? string.Empty)));
                list.Add(item);
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
            var list = new List<EmployeeHasHeader>();
            foreach (var item in Purcell.Query<EmployeeHasHeader>(file, "Sheet1"))
            {
                Assert.True(item.EmpId > 0);
                Assert.True(Convert.ToString(item.RealName ?? string.Empty).StartsWith(Convert.ToString(item.LastName ?? string.Empty)));
                list.Add(item);
            }

            Assert.Equal(10, list.Count);
        }
    }
}