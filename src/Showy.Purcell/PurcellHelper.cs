using System;
using System.Text;

namespace Showy.Purcell;

/// <summary>
/// Purcell帮助类
/// </summary>
public class PurcellHelper
{
    /// <summary>
    /// 按照Excel规则转换为DateTime，<br />
    /// 注意：如果是Excel for Mac 2008及更早版本，请自行在结果上减去1462天。<br />
    /// https://support.microsoft.com/en-us/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487
    /// </summary>
    /// <param name="value"></param>
    /// <param name="dt"></param>
    /// <returns></returns>
    public static bool TryParseDateTime(double value, out DateTime dt)
    {
        dt = DateTime.MinValue;
        var dateTime = PurcellConst.Epoch1900;

        switch (value)
        {
            case < 0.0:
                return false;
            case < 1.0:
                // 这里实际返回了0000-01-01年
                dt = DateTime.MinValue.AddDays(value);
                return true;
            case < 61.0 and >= 60.0:
                // 1900年不是闰年，但是Excel错误的认为是闰年，所以1900-02-29返回false
                return false;
            case < 61.0:
                value += 1.0;
                break;
        }

        dt = dateTime.AddDays(value);
        return true;
    }

    /// <summary>
    /// 仅用于数值字符串
    /// </summary>
    /// <param name="value"></param>
    /// <param name="dt"></param>
    /// <returns></returns>
    internal static bool TryParseDateTime(string value, out DateTime dt)
    {
        dt = DateTime.MinValue;
        return double.TryParse(value, out var valueDouble) && TryParseDateTime(valueDouble, out dt);
    }

    /// <summary>
    /// 清除空白字符
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    internal static string? ClearWhiteSpace(string? value)
    {
        if (value == null) return value;

        var sb = new StringBuilder();

        foreach (var c in value)
        {
            if (!char.IsWhiteSpace(c))
            {
                sb.Append(c);
            }
        }

        return sb.ToString();
    }
}