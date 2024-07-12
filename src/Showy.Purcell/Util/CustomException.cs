using System;
using System.Collections.Generic;
using System.Text;

namespace Showy.Purcell;

/// <summary>
/// 重复的表头异常
/// </summary>
public class DuplicateHeaderException : Exception
{
    /// <inheritdoc cref="DuplicateHeaderException"/>
    /// <param name="message"></param>
    public DuplicateHeaderException(string message) : base(message)
    {
    }
}