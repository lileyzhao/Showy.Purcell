using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;
using Showy.Purcell.Tests.Dto;

namespace Showy.Purcell.Tests;

public class QueryOthers : QueryBase
{
    [Fact]
    public void QueryGenericDateTime()
    {
        var epoch = new DateTime(1900, 1, 1);
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryGeneric)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var i = -1;
            foreach (var item in Purcell.Query<DateTimeDto>(file, (int)SheetIndex.Sheet1).ToList())
            {
                i++;
                Debug.WriteLine("当前行: " + i);
                Assert.Equal(i / 10d, Convert.ToDouble(item.Value));
                
                Assert.Equal(DateTime.MinValue, item.Date);
                Assert.Equal(DateTime.MinValue, item.DateTime);
                Assert.Equal(DateTime.MinValue.AddMinutes(i * 144), item.Time);
                
                Assert.Equal(DateTime.MinValue, item.TimeMin);
                Assert.Null(item.TimeNull);
            }
        }
    }

    [Fact]
    public void QueryGenericEnum()
    {
        var fileName = Regex.Match(MethodBase.GetCurrentMethod()!.Name, "(?<=QueryGeneric)[^_]+").Value;
        foreach (var file in GetFiles(fileName))
        {
            var i = -1;
            foreach (var item in Purcell.Query<EnumDto>(file, SheetIndex.Sheet1))
            {
                i++;
                if (i < 10)
                {
                    Assert.Equal(i % 2 == 0 ? Sex.Man : Sex.Woman, item.Sex);
                }
                else
                {
                    Assert.Equal(0, (int)item.Sex);
                }
            }
        }
    }
}