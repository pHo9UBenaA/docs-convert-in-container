using System;
using System.Threading;
using System.Threading.Tasks;

namespace SharedXmlToJsonl.ErrorHandling;

public interface IGlobalErrorHandler
{
    Task<int> HandleErrorAsync(Exception exception, CancellationToken cancellationToken = default);
    void LogError(Exception exception, string context);
    string FormatErrorMessage(Exception exception);
}