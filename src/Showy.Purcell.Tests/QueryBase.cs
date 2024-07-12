using System.Runtime.CompilerServices;
using System.Text;

namespace Showy.Purcell.Tests;

public abstract class QueryBase
{
    private const string FileFormat = "Data/{0}";
    private static readonly string[] TestExtList = [".csv", ".xls", ".xlsx"];
    protected static readonly ExcelFileType[] TestExtListEnum = [ExcelFileType.Csv, ExcelFileType.Xls, ExcelFileType.Xlsx];

    protected virtual string[] GetFiles([CallerMemberName] string name = "")
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var files = new string[TestExtList.Length];
        for (var i = 0; i < TestExtList.Length; i++)
        {
            var file = string.Format(FileFormat, name) + TestExtList[i];

            Assert.True(File.Exists(file), "测试用的Excel文件： " + file + " 不存在或无法读取。");
            files[i] = file;
        }

        return files;
    }

    protected virtual Stream[] GetStreams([CallerMemberName] string name = "")
    {
        var streams = new Stream[TestExtList.Length];
        for (var i = 0; i < TestExtList.Length; i++)
        {
            var file = string.Format(FileFormat, name) + TestExtList[i];

            Assert.True(File.Exists(file), "测试用的Excel文件： " + file + " 不存在或无法读取。");
            streams[i] = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read, 1);
        }

        return streams;
    }

    public static IDictionary<string, object?> ToDictionary(dynamic dynamicObject)
    {
        IDictionary<string, object?> dictionary = dynamicObject;

        return dictionary;
    }
}