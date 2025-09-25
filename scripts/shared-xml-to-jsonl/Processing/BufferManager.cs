using System;
using System.Buffers;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Processing;

public partial class BufferManager : IBufferManager
{
    private const int DefaultBufferSize = 4096;
    private readonly ILogger<BufferManager> _logger;
    private bool _disposed;

    [LoggerMessage(EventId = 1, Level = LogLevel.Trace, Message = "Rented byte buffer of size {Size}")]
    private static partial void LogByteBufferRented(ILogger logger, int size);

    [LoggerMessage(EventId = 2, Level = LogLevel.Trace, Message = "Returned byte buffer of size {Size}")]
    private static partial void LogByteBufferReturned(ILogger logger, int size);

    [LoggerMessage(EventId = 3, Level = LogLevel.Trace, Message = "Rented char buffer of size {Size}")]
    private static partial void LogCharBufferRented(ILogger logger, int size);

    [LoggerMessage(EventId = 4, Level = LogLevel.Trace, Message = "Returned char buffer of size {Size}")]
    private static partial void LogCharBufferReturned(ILogger logger, int size);

    [LoggerMessage(EventId = 5, Level = LogLevel.Debug, Message = "BufferManager disposed")]
    private static partial void LogBufferManagerDisposed(ILogger logger);

    public BufferManager(ILogger<BufferManager> logger)
    {
        ArgumentNullException.ThrowIfNull(logger);
        _logger = logger;
    }

    public byte[] RentBuffer(int minimumSize = DefaultBufferSize)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(BufferManager));

        var buffer = ArrayPool<byte>.Shared.Rent(minimumSize);
        LogByteBufferRented(_logger, buffer.Length);
        return buffer;
    }

    public void ReturnBuffer(byte[] buffer)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(BufferManager));

        ArgumentNullException.ThrowIfNull(buffer);

        ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
        LogByteBufferReturned(_logger, buffer.Length);
    }

    public char[] RentCharBuffer(int minimumSize = DefaultBufferSize)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(BufferManager));

        var buffer = ArrayPool<char>.Shared.Rent(minimumSize);
        LogCharBufferRented(_logger, buffer.Length);
        return buffer;
    }

    public void ReturnCharBuffer(char[] buffer)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(BufferManager));

        ArgumentNullException.ThrowIfNull(buffer);

        ArrayPool<char>.Shared.Return(buffer, clearArray: true);
        LogCharBufferReturned(_logger, buffer.Length);
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                // Clean up managed resources if any
                LogBufferManagerDisposed(_logger);
            }
            _disposed = true;
        }
    }
}
