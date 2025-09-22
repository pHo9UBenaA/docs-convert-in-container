using System;
using System.Threading;
using System.Threading.Tasks;
using Xunit;
using Moq;
using FluentAssertions;
using Microsoft.Extensions.Logging;
using PptxXmlToJsonl.Commands;
using SharedXmlToJsonl;
using SharedXmlToJsonl.Interfaces;
using SharedXmlToJsonl.Models;

namespace PptxXmlToJsonl.Tests.Commands;

public class ConvertPptxCommandTests
{
    private readonly Mock<ILogger<ConvertPptxCommand>> _loggerMock;
    private readonly Mock<IPptxProcessor> _processorMock;
    private readonly Mock<IJsonWriter> _jsonWriterMock;
    private readonly ConvertPptxCommand _sut;

    public ConvertPptxCommandTests()
    {
        _loggerMock = new Mock<ILogger<ConvertPptxCommand>>();
        _processorMock = new Mock<IPptxProcessor>();
        _jsonWriterMock = new Mock<IJsonWriter>();

        _sut = new ConvertPptxCommand(
            _loggerMock.Object,
            Mock.Of<IServiceProvider>(),
            _processorMock.Object,
            _jsonWriterMock.Object);
    }

    [Fact]
    public async Task ExecuteAsync_ValidInput_ReturnsSuccess()
    {
        // Arrange
        var options = new ConvertPptxOptions
        {
            InputPath = "test.pptx",
            OutputDirectory = "output"
        };

        _processorMock
            .Setup(x => x.ProcessAsync(
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<CancellationToken>()))
            .ReturnsAsync(new ProcessingResult { Success = true });

        // Act
        var result = await _sut.ExecuteAsync(options);

        // Assert
        result.Should().Be(CommonBase.ExitSuccess);
        _processorMock.Verify(x => x.ProcessAsync(
            "test.pptx",
            "output",
            It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task ExecuteAsync_FileNotFound_ReturnsError()
    {
        // Arrange
        var options = new ConvertPptxOptions
        {
            InputPath = "nonexistent.pptx",
            OutputDirectory = "output"
        };

        _processorMock
            .Setup(x => x.ProcessAsync(
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<CancellationToken>()))
            .ThrowsAsync(new FileNotFoundException("File not found"));

        // Act
        var result = await _sut.ExecuteAsync(options);

        // Assert
        result.Should().Be(CommonBase.ExitProcessingError);
        _loggerMock.Verify(
            x => x.Log(
                LogLevel.Error,
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((v, t) => v.ToString().Contains("Command execution failed")),
                It.IsAny<Exception>(),
                It.IsAny<Func<It.IsAnyType, Exception, string>>()),
            Times.Once);
    }

    [Fact]
    public async Task ExecuteAsync_InvalidOptions_ReturnsUsageError()
    {
        // Arrange
        var options = new ConvertPptxOptions
        {
            InputPath = "",  // Invalid: empty path
            OutputDirectory = "output"
        };

        // Act
        var result = await _sut.ExecuteAsync(options);

        // Assert
        result.Should().Be(CommonBase.ExitUsageError);
        _processorMock.Verify(x => x.ProcessAsync(
            It.IsAny<string>(),
            It.IsAny<string>(),
            It.IsAny<CancellationToken>()), Times.Never);
    }

    [Fact]
    public async Task ExecuteAsync_CancellationRequested_ReturnsProcessingError()
    {
        // Arrange
        var options = new ConvertPptxOptions
        {
            InputPath = "test.pptx",
            OutputDirectory = "output"
        };

        var cts = new CancellationTokenSource();
        cts.Cancel();

        _processorMock
            .Setup(x => x.ProcessAsync(
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<CancellationToken>()))
            .ThrowsAsync(new OperationCanceledException());

        // Act
        var result = await _sut.ExecuteAsync(options, cts.Token);

        // Assert
        result.Should().Be(CommonBase.ExitProcessingError);
        _loggerMock.Verify(
            x => x.Log(
                LogLevel.Warning,
                It.IsAny<EventId>(),
                It.Is<It.IsAnyType>((v, t) => v.ToString().Contains("Operation was cancelled")),
                It.IsAny<Exception>(),
                It.IsAny<Func<It.IsAnyType, Exception, string>>()),
            Times.Once);
    }

    [Fact]
    public async Task ExecuteAsync_ProcessorReturnsFailure_ReturnsError()
    {
        // Arrange
        var options = new ConvertPptxOptions
        {
            InputPath = "test.pptx",
            OutputDirectory = "output"
        };

        _processorMock
            .Setup(x => x.ProcessAsync(
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<CancellationToken>()))
            .ReturnsAsync(new ProcessingResult
            {
                Success = false,
                ErrorMessage = "Processing failed"
            });

        // Act
        var result = await _sut.ExecuteAsync(options);

        // Assert
        result.Should().Be(CommonBase.ExitProcessingError);
    }
}