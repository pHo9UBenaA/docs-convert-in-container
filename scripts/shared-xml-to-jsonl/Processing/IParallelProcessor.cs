using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace SharedXmlToJsonl.Processing;

public interface IParallelProcessor
{
    Task ProcessBatchAsync<T>(
        IEnumerable<T> items,
        Func<T, CancellationToken, Task> processItem,
        CancellationToken cancellationToken = default);

    Task ProcessBatchAsync<T, TResult>(
        IEnumerable<T> items,
        Func<T, CancellationToken, Task<TResult>> processItem,
        Action<TResult> resultHandler,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<TResult>> ProcessBatchAsync<T, TResult>(
        IEnumerable<T> items,
        Func<T, CancellationToken, Task<TResult>> processItem,
        CancellationToken cancellationToken = default);
}
