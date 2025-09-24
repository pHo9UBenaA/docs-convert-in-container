using System;
using System.Buffers;
using Microsoft.Extensions.Logging;

namespace SharedXmlToJsonl.Processing;

public class BufferManager : IBufferManager
{
    private const int DefaultBufferSize = 4096;
    private readonly ILogger<BufferManager> _logger;
    private bool _disposed;

    public BufferManager(ILogger<BufferManager> logger)
    {
        ArgumentNullException.ThrowIfNull(logger);
        _logger = logger;
    }

    public byte[] RentBuffer(int minimumSize = DefaultBufferSize)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(BufferManager));

        var buffer = ArrayPool<byte>.Shared.Rent(minimumSize);
        _logger.LogTrace("Rented byte buffer of size {Size}", buffer.Length);
        return buffer;
    }

    public void ReturnBuffer(byte[] buffer)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(BufferManager));

        ArgumentNullException.ThrowIfNull(buffer);

        ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
        _logger.LogTrace("Returned byte buffer of size {Size}", buffer.Length);
    }

    public char[] RentCharBuffer(int minimumSize = DefaultBufferSize)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(BufferManager));

        var buffer = ArrayPool<char>.Shared.Rent(minimumSize);
        _logger.LogTrace("Rented char buffer of size {Size}", buffer.Length);
        return buffer;
    }

    public void ReturnCharBuffer(char[] buffer)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(BufferManager));

        ArgumentNullException.ThrowIfNull(buffer);

        ArrayPool<char>.Shared.Return(buffer, clearArray: true);
        _logger.LogTrace("Returned char buffer of size {Size}", buffer.Length);
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
                _logger.LogDebug("BufferManager disposed");
            }
            _disposed = true;
        }
    }
}
