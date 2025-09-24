using System;
using System.ComponentModel.DataAnnotations;

namespace SharedXmlToJsonl.Configuration;

public class ProcessingOptions
{
    [Range(1, 10000)]
    public int MaxConcurrency { get; set; } = 10;

    [Required]
    public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(5);

    public bool EnableCaching { get; set; } = true;

    [Range(1, 100)]
    public int CacheSizeMB { get; set; } = 50;

    public RetryOptions Retry { get; set; } = new();

    public bool Verbose { get; set; }

    public bool IncludeHiddenSlides { get; set; }

    public bool ExtractShapes { get; set; } = true;

    public bool ExtractText { get; set; } = true;

    [Range(1, 1000)]
    public int MaxSlides { get; set; } = 1000;

    [Range(1, 10000)]
    public int MaxSheets { get; set; } = 1000;

    public string? SheetRange { get; set; }

    public bool IncludeFormulas { get; set; } = true;

    public bool IncludeEmptyCells { get; set; }
}

public class RetryOptions
{
    [Range(0, 10)]
    public int MaxAttempts { get; set; } = 3;

    public TimeSpan InitialDelay { get; set; } = TimeSpan.FromSeconds(1);

    public double BackoffMultiplier { get; set; } = 2.0;

    public TimeSpan MaxDelay { get; set; } = TimeSpan.FromSeconds(30);
}
