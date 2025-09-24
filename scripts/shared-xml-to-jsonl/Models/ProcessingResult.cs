using System;
using System.Collections.Generic;

namespace SharedXmlToJsonl.Models;

public class ProcessingResult
{
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<string> Errors { get; set; } = new List<string>();
    public int ItemsProcessed { get; set; }
    public TimeSpan ElapsedTime { get; set; }
    public string? OutputPath { get; set; }
    public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();
}
