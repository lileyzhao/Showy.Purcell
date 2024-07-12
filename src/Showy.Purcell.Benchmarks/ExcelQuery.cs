using BenchmarkDotNet.Attributes;
using ExcelDataReader;
using MiniExcelLibs;
using SharpCompress.Common;
using Sylvan.Data.Excel;

namespace Showy.Purcell.Benchmarks;

[MemoryDiagnoser]
[ShortRunJob]
[HtmlExporter]
public class ExcelQuery
{
    [GlobalSetup]
    public void Setup()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }

    private const string FileName = @"Data/10_000x11.xlsx";

    [Benchmark(Description = "Purcell_Query_Dynamic")]
    public void Purcell_Query_Dynamic()
    {
        var path = Path.Combine(AppContext.BaseDirectory, FileName);
        foreach (var unused in Purcell.Query(path, 0, -1, 0))
        {
        }
    }

    [Benchmark(Description = "Sylvan_Query_Dynamic")]
    public void Sylvan_Query_Dynamic()
    {
        var path = Path.Combine(AppContext.BaseDirectory, FileName);
        using var edr = Sylvan.Data.Excel.ExcelDataReader.Create(path);
        do
        {
            while (edr.Read())
                for (var i = 0; i < edr.FieldCount; i++)
                {
                    var unused = edr.GetString(i);
                }
        } while (edr.NextResult());
    }

    [Benchmark(Description = "MiniExcel_Query_Dynamic")]
    public void MiniExcel_Query_Dynamic()
    {
        var path = Path.Combine(AppContext.BaseDirectory, FileName);
        foreach (var unused in MiniExcel.Query(path))
        {
        }
    }


    //[Benchmark(Description = "ExcelDataReader_Query_Dynamic")]
    public void ExcelDataReader_Query_Dynamic()
    {
        var path = Path.Combine(AppContext.BaseDirectory, FileName);
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read);
        using var edr = ExcelReaderFactory.CreateReader(stream);
        do
        {
            while (edr.Read())
            {
                for (var i = 0; i < edr.FieldCount; i++)
                {
                    var unused = edr.GetValue(i);
                }
            }
        } while (edr.NextResult());
    }

    // [Benchmark(Description = "Purcell_Query_Dynamic_First")]
    public void Purcell_Query_Dynamic_First()
    {
        var path = Path.Combine(AppContext.BaseDirectory, FileName);
        var unused = Purcell.Query(path, 0, -1, 0).First();
    }

    // [Benchmark(Description = "MiniExcel_Query_Dynamic_First")]
    public void MiniExcel_Query_Dynamic_First()
    {
        var path = Path.Combine(AppContext.BaseDirectory, FileName);
        var unused = MiniExcel.Query(path).First();
    }

    //[Benchmark(Description = "Purcell_Query_Generic")]
    public void Purcell_Query_Generic()
    {
        //var path = Path.Combine(AppContext.BaseDirectory, fileName);
        //foreach (var unused in Showy.Purcell.Purcell.Query<TestModel>(path, 0, headerRow: -1, dataStartRow: 0))
        //{
        //}
    }

    //[Benchmark(Description = "MiniExcel_Query_Generic")]
    public void MiniExcel_Query_Generic()
    {
        //var path = Path.Combine(AppContext.BaseDirectory, fileName);
        //foreach (var unused in MiniExcelLibs.MiniExcel.Query<TestModel>(path))
        //{
        //}
    }
}