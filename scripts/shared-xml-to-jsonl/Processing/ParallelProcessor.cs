using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SharedXmlToJsonl.Configuration;

namespace SharedXmlToJsonl.Processing;

public class ParallelProcessor : IParallelProcessor
{
    private readonly ILogger<ParallelProcessor> _logger;
    private readonly ProcessingOptions _options;

    public ParallelProcessor(
        ILogger<ParallelProcessor> logger,
        IOptions<ProcessingOptions> options)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _options = options?.Value ?? throw new ArgumentNullException(nameof(options));
    }

    public async Task ProcessBatchAsync<T>(
        IEnumerable<T> items,
        Func<T, CancellationToken, Task> processItem,
        CancellationToken cancellationToken = default)
    {
        if (items == null) throw new ArgumentNullException(nameof(items));
        if (processItem == null) throw new ArgumentNullException(nameof(processItem));

        var itemList = items.ToList();
        if (!itemList.Any()) return;

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
                    _logger.LogError(ex, "Error processing item: {Item}", item);
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
        if (items == null) throw new ArgumentNullException(nameof(items));
        if (processItem == null) throw new ArgumentNullException(nameof(processItem));
        if (resultHandler == null) throw new ArgumentNullException(nameof(resultHandler));

        var itemList = items.ToList();
        if (!itemList.Any()) return;

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
                    _logger.LogError(ex, "Error processing item: {Item}", item);
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
        if (items == null) throw new ArgumentNullException(nameof(items));
        if (processItem == null) throw new ArgumentNullException(nameof(processItem));

        var itemList = items.ToList();
        if (!itemList.Any()) return Array.Empty<TResult>();

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
                    _logger.LogError(ex, "Error processing item: {Item}", item);
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
}