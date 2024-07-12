using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Showy.Purcell;

/// <summary>
/// Type扩展类 [Extensions]
/// </summary>
internal static class TypeExtensions
{
    /// <summary>
    /// 获取程序集属性
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="assembly"></param>
    /// <param name="inherit"></param>
    /// <returns></returns>
    public static T? GetAttribute<T>(this ICustomAttributeProvider? assembly, bool inherit = false)
        where T : Attribute
    {
        return assembly?.GetCustomAttributes(typeof(T), inherit).OfType<T>().FirstOrDefault();
    }

    /// <summary>
    /// 获取显示名
    /// </summary>
    /// <param name="assembly"></param>
    /// <param name="inherit"></param>
    /// <returns></returns>
    public static string? GetDisplayName(this ICustomAttributeProvider? assembly, bool inherit = false)
    {
        // DisplayAttribute? displayAttribute = assembly.GetAttribute<DisplayAttribute>(inherit);
        // return displayAttribute != null ? displayAttribute.Name : assembly.GetAttribute<DisplayNameAttribute>()?.DisplayName;
        return assembly.GetAttribute<DisplayNameAttribute>()?.DisplayName;
    }

    /// <summary>
    /// 获取显示名
    /// </summary>
    /// <param name="assembly"></param>
    /// <param name="inherit"></param>
    /// <returns></returns>
    public static string? GetDescription(this ICustomAttributeProvider? assembly, bool inherit = false)
    {
        return assembly.GetAttribute<DescriptionAttribute>()?.Description;
    }

    /// <summary>
    /// 获取类型名称
    /// <para>Ps. 额外处理泛型类型</para>
    /// </summary>
    /// <param name="type"></param>
    /// <returns></returns>
    public static string GetCSharpTypeName(this Type type)
    {
        StringBuilder sb = new();
        var name = type.Name;
        if (!type.IsGenericType) return name;
        sb.Append(name.Substring(0, name.IndexOf('`')));
        sb.Append('<');
        sb.Append(string.Join(", ", type.GetGenericArguments().Select(t => t.GetCSharpTypeName())));

        sb.Append('>');
        return sb.ToString();
    }

    // 辅助方法用于将 IAsyncEnumerable 转换为 IEnumerable
    internal static IEnumerable<T> ToEnumerable<T>(this IAsyncEnumerable<T> asyncEnumerable)
    {
        // 注意：此方法会阻塞调用线程，直到异步操作完成
        var enumerator = asyncEnumerable.GetAsyncEnumerator();
        while (enumerator.MoveNextAsync().AsTask().Result)
        {
            yield return enumerator.Current;
        }
    }
}