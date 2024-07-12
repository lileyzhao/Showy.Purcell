using System.Diagnostics;
using MiniExcelLibs;
using Showy.Purcell;
using Showy.Purcell.Consoles;
using Sylvan.Data.Excel;

const string testFile = @"Data/10_000x11.xlsx";

/*
 * 控制台项目只是用于随手测试
 */
Console.WriteLine("Hello, Purcell from the Showy Team!");

//var currentProcess = Process.GetCurrentProcess();

//// Purcell
//foreach (var unused in Purcell.Query(Path.Combine(AppContext.BaseDirectory, testFile)))
//{
//}

//// MiniExcel
//foreach (var unused in MiniExcel.Query(Path.Combine(AppContext.BaseDirectory, testFile)).Cast<IDictionary<string, object?>>())
//{
//}

//// Sylvan.Data.Excel
//using var edr = ExcelDataReader.Create(Path.Combine(AppContext.BaseDirectory, testFile));
//edr.NextResult();
//while (edr.Read())
//    for (var i = 0; i < edr.FieldCount; i++)
//    {
//        var unused = edr.GetString(i);
//    }

//Console.WriteLine("测试完成!");
//var peakMemory = currentProcess.PeakWorkingSet64;

//Console.WriteLine($"运行期间的峰值内存占用量: {peakMemory / (1024.0 * 1024.0)} MB");
//Console.ReadLine();


//UsaAddress.Proc();