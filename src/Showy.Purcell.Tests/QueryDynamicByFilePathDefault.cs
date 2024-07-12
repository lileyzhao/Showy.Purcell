using System.Reflection;
using System.Text.RegularExpressions;

namespace Showy.Purcell.Tests;

public class QueryDynamicByFilePathDefault : QueryBase
{
    [Fact]
    public void QueryDynamicHasHeader_OR1()
    {
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryDynamic)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var list = new List<IDictionary<string, object?>>();
            foreach (var item in Purcell.Query(file))
            {
                Assert.True(int.TryParse(item["员工ID"]!.ToString(), out var _));
                Assert.True(Convert.ToString(item["姓名"])!.StartsWith(Convert.ToString(item["姓氏"])!));
                list.Add(item);
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
            var list = new List<IDictionary<string, object?>>();
            foreach (var item in Purcell.Query(file, 0))
            {
                Assert.True(int.TryParse(item["员工ID"]!.ToString(), out var _));
                Assert.True(Convert.ToString(item["姓名"])!.StartsWith(Convert.ToString(item["姓氏"])!));
                list.Add(item);
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
            var list = new List<IDictionary<string, object?>>();
            foreach (var item in Purcell.Query(file, SheetIndex.Sheet1))
            {
                Assert.True(int.TryParse(item["员工ID"]!.ToString(), out var _));
                Assert.True(Convert.ToString(item["姓名"])!.StartsWith(Convert.ToString(item["姓氏"])!));
                list.Add(item);
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
            var list = new List<IDictionary<string, object?>>();
            foreach (var item in Purcell.Query(file, "Sheet1"))
            {
                Assert.True(int.TryParse(item["员工ID"]!.ToString(), out var _));
                Assert.True(Convert.ToString(item["姓名"])!.StartsWith(Convert.ToString(item["姓氏"])!));
                list.Add(item);
            }

            Assert.Equal(10, list.Count);
        }
    }
}