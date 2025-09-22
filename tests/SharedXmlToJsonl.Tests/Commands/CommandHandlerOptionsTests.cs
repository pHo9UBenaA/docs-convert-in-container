using System.IO;
using Xunit;
using FluentAssertions;
using SharedXmlToJsonl.Commands;

namespace SharedXmlToJsonl.Tests.Commands;

public class CommandHandlerOptionsTests
{
    private class TestOptions : CommandHandlerOptions
    {
        public string? CustomProperty { get; set; }
    }

    [Fact]
    public void Validate_EmptyInputPath_ReturnsInvalid()
    {
        // Arrange
        var options = new TestOptions
        {
            InputPath = "",
            OutputDirectory = "output"
        };

        // Act
        var result = options.Validate();

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.Contains("Input path is required"));
    }

    [Fact]
    public void Validate_EmptyOutputDirectory_ReturnsInvalid()
    {
        // Arrange
        var options = new TestOptions
        {
            InputPath = "test.pptx",
            OutputDirectory = ""
        };

        // Act
        var result = options.Validate();

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.Contains("Output directory is required"));
    }

    [Fact]
    public void Validate_ValidOptions_ReturnsValid()
    {
        // Arrange
        var options = new TestOptions
        {
            InputPath = "test.pptx",
            OutputDirectory = "output",
            Verbose = true
        };

        // Act
        var result = options.Validate();

        // Assert
        result.IsValid.Should().BeTrue();
        result.Errors.Should().BeEmpty();
    }

    [Fact]
    public void Validate_BothInputAndOutputEmpty_ReturnsMultipleErrors()
    {
        // Arrange
        var options = new TestOptions
        {
            InputPath = "",
            OutputDirectory = ""
        };

        // Act
        var result = options.Validate();

        // Assert
        result.IsValid.Should().BeFalse();
        result.Errors.Should().HaveCount(2);
        result.Errors.Should().Contain(e => e.Contains("Input path is required"));
        result.Errors.Should().Contain(e => e.Contains("Output directory is required"));
    }
}