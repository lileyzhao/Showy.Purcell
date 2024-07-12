// See https://aka.ms/new-console-template for more information

using BenchmarkDotNet.Running;
using Showy.Purcell.Benchmarks;

#if DEBUG

new Showy.Purcell.Benchmarks.ExcelQuery().Sylvan_Query_Dynamic();
// 输入任意字符串继续
Console.WriteLine("Press any key to continue...");
Console.ReadKey();

#else

var summary = BenchmarkRunner.Run<ExcelQuery>();

#endif