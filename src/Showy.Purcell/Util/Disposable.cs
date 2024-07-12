using System;

namespace Showy.Purcell;

/// <summary>
/// 通用Disposable
/// </summary>
internal abstract class Disposable : IDisposable
{
    private bool _isDisposed;

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// 终结器 Finalizer
    /// </summary>
    ~Disposable()
    {
        // cleanup statements...
        // 清理语句...
        Dispose(false);
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    /// <param name="disposing"></param>
    private void Dispose(bool disposing)
    {
        if (!_isDisposed && disposing) DisposeCore();

        _isDisposed = true;
    }

    /// <summary>
    /// 释放资源，由继承者实现
    /// </summary>
    protected abstract void DisposeCore();
}