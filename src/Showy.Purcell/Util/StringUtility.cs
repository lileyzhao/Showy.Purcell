using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Showy.Purcell;

/// <summary>
/// 字符串工具类 String utility class
/// </summary>
internal class StringUtility
{
    /// <summary>
    /// 判断字符是否是全角字符 Determine whether the character is a full-width character
    /// </summary>
    /// <param name="c">字符 character</param>
    public static bool IsFullWidth(char c)
    {
        // 判断是否为全角形式字符 Determine whether it is a full-width character
        if (c >= 0xFF01 && c <= 0xFF5E) // 全角ASCII码范围 Full-width ASCII code range
        {
            return true;
        }

        // 判断是否为CJK统一表意符号，常用于中文、日文和韩文 Determine whether it is a CJK unified ideograph symbol, commonly used in Chinese, Japanese, and Korean
        if (c >= 0x4E00 && c <= 0x9FFF || // 基本区块 Basic block
            c >= 0x3400 && c <= 0x4DBF || // 扩展A区块 Extended A block
            c >= 0x20000 && c <= 0x2A6DF || // 扩展B区块 Extended B block
            c >= 0x2A700 && c <= 0x2B73F || // 扩展C区块 Extended C block
            c >= 0x2B740 && c <= 0x2B81F || // 扩展D区块 Extended D block
            c >= 0x2B820 && c <= 0x2CEAF || // 扩展E区块 Extended E block
            c >= 0xF900 && c <= 0xFAFF || // 兼容表意符号 Compatible ideograph symbol
            c >= 0x2F800 && c <= 0x2FA1F)  // 兼容扩展 Compatible extension
        {
            return true;
        }

        // 判断是否为韩文全角形式 Determine whether it is a full-width Korean form
        if (c >= 0xAC00 && c <= 0xD7A3) // 韩文字符集 Korean character set
        {
            return true;
        }

        // 判断是否为日文全角形式，包含假名和片假名 Determine whether it is a full-width Japanese form, including hiragana and katakana
        if ((c >= 0x3040 && c <= 0x309F) || // 平假名 Hiragana
            (c >= 0x30A0 && c <= 0x30FF) || // 片假名 Katakana
            (c >= 0x31F0 && c <= 0x31FF))   // 片假名扩展 Katakana extension
        {
            return true;
        }

        // 判断是否为全角标点符号 Determine whether it is a full-width punctuation mark
        if (c >= 0x3000 && c <= 0x303F)
        {
            return true;
        }

        return false; // 默认返回false Default return false
    }
}
