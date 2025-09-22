using System;

namespace SharedXmlToJsonl.Resources;

public interface IResourceService
{
    string GetErrorMessage(string key, params object[] args);
    string GetLogMessage(string key, params object[] args);
    string GetString(string resourceName, string key, params object[] args);
    bool TryGetString(string resourceName, string key, out string? value);
}