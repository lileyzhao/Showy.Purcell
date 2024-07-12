using System;
using System.Threading;
using System.Resources;

namespace Showy.Purcell;

/// <summary>
/// 多语言工具类 Multilingual utility class
/// </summary>
internal static class Multilingual
{
    public static T CreateException<T>(string errorCode, Language? language = null) where T : Exception
    {
        if (string.IsNullOrWhiteSpace(errorCode))
            throw new ArgumentNullException(nameof(errorCode), "errorCode cannot be blank. (errorCode不允许为空白字符)");

        // 根据语言设置当前线程的语言环境
        Thread.CurrentThread.CurrentUICulture = language switch
        {
            Language.English => new System.Globalization.CultureInfo("en-US"),
            Language.Chinese => new System.Globalization.CultureInfo("zh-CN"),
            _ => Thread.CurrentThread.CurrentUICulture
        };

        // 读取资源文件
        ResourceManager rm = new ResourceManager("YourNamespace.ErrorMessages", typeof(Multilingual).Assembly);

        // 获取对应语言的错误消息，如果获取为空，则直接使用 errorCode
        var errorMessage = rm.GetString(errorCode, Thread.CurrentThread.CurrentUICulture) ?? errorCode;

        // 抛出异常
        return (T?)Activator.CreateInstance(typeof(T), errorMessage)
            ?? (T)new Exception("Error occurred when throwing an exception. (抛出异常时发生错误)");
    }

    public static Exception CreateException(string errorCode, Language? language = null)
    {
        return CreateException<Exception>(errorCode, language);
    }
}