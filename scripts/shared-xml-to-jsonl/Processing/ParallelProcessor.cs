using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SharedXmlToJsonl.Configuration;

namespace SharedXmlToJsonl.Processing;

public partial class ParallelProcessor : IParallelProcessor
{
    private readonly ILogger<ParallelProcessor> _logger;
    private readonly ProcessingOptions _options;

    public ParallelProcessor(
        ILogger<ParallelProcessor> logger,
        IOptions<ProcessingOptions> options)
    {
        ArgumentNullException.ThrowIfNull(logger);
        ArgumentNullException.ThrowIfNull(options);
        _logger = logger;
        _options = options.Value;
    }

    public async Task ProcessBatchAsync<T>(
        IEnumerable<T> items,
        Func<T, CancellationToken, Task> processItem,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(items);
        ArgumentNullException.ThrowIfNull(processItem);

        var itemList = items.ToList();
        if (itemList.Count == 0) return;

        using var semaphore = new SemaphoreSlim(_options.MaxConcurrency);
        var tasks = new List<Task>();

        foreach (var item in itemList)
        {
            await semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);

            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    await processItem(item, cancellationToken).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    LogErrorProcessingItem(_logger, ex, item?.ToString() ?? "null");
                    throw;
                }
                finally
                {
                    semaphore.Release();
                }
            }, cancellationToken));
        }

        await Task.WhenAll(tasks).ConfigureAwait(false);
    }

    public async Task ProcessBatchAsync<T, TResult>(
        IEnumerable<T> items,
        Func<T, CancellationToken, Task<TResult>> processItem,
        Action<TResult> resultHandler,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(items);
        ArgumentNullException.ThrowIfNull(processItem);
        ArgumentNullException.ThrowIfNull(resultHandler);

        var itemList = items.ToList();
        if (itemList.Count == 0) return;

        using var semaphore = new SemaphoreSlim(_options.MaxConcurrency);
        var tasks = new List<Task>();
        var resultLock = new object();

        foreach (var item in itemList)
        {
            await semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);

            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    var result = await processItem(item, cancellationToken).ConfigureAwait(false);
                    lock (resultLock)
                    {
                        resultHandler(result);
                    }
                }
                catch (Exception ex)
                {
                    LogErrorProcessingItem(_logger, ex, item?.ToString() ?? "null");
                    throw;
                }
                finally
                {
                    semaphore.Release();
                }
            }, cancellationToken));
        }

        await Task.WhenAll(tasks).ConfigureAwait(false);
    }

    public async Task<IReadOnlyList<TResult>> ProcessBatchAsync<T, TResult>(
        IEnumerable<T> items,
        Func<T, CancellationToken, Task<TResult>> processItem,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(items);
        ArgumentNullException.ThrowIfNull(processItem);

        var itemList = items.ToList();
        if (itemList.Count == 0) return Array.Empty<TResult>();

        using var semaphore = new SemaphoreSlim(_options.MaxConcurrency);
        var tasks = new List<Task<TResult>>();

        foreach (var item in itemList)
        {
            await semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);

            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    var result = await processItem(item, cancellationToken).ConfigureAwait(false);
                    return result;
                }
                catch (Exception ex)
                {
                    LogErrorProcessingItem(_logger, ex, item?.ToString() ?? "null");
                    throw;
                }
                finally
                {
                    semaphore.Release();
                }
            }, cancellationToken));
        }

        var results = await Task.WhenAll(tasks).ConfigureAwait(false);
        return results;
    }

    [LoggerMessage(
        EventId = 7001,
        Level = LogLevel.Error,
        Message = "Error processing item: {item}")]
    private static partial void LogErrorProcessingItem(
        ILogger logger, Exception ex, string item);
}
