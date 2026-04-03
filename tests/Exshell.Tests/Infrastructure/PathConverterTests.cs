using Exshell.Infrastructure;

namespace Exshell.Tests.Infrastructure;

public class PathConverterTests
{
    [Fact]
    public void NormalizeWindowsPath_WithWindowsPath_ReturnsNormalizedPath()
    {
        // Arrange
        var input = @"D:\work\sample.xlsx";

        // Act
        var result = PathConverter.NormalizeWindowsPath(input);

        // Assert
        Assert.NotNull(result);
        Assert.Contains(@"D:\work\sample.xlsx", result);
    }

    [Fact]
    public void NormalizeWindowsPath_WithWslMountPath_ReturnsWindowsPath()
    {
        // Arrange
        var input = "/mnt/d/work/sample.xlsx";

        // Act
        var result = PathConverter.NormalizeWindowsPath(input);

        // Assert
        Assert.NotNull(result);
        Assert.Contains(@"D:\work\sample.xlsx", result);
    }

    [Fact]
    public void NormalizeWindowsPath_WithQuotedPath_RemovesQuotes()
    {
        // Arrange
        var input = @"""D:\work\sample.xlsx""";

        // Act
        var result = PathConverter.NormalizeWindowsPath(input);

        // Assert
        Assert.NotNull(result);
        Assert.DoesNotContain("\"", result);
        Assert.Contains(@"D:\work\sample.xlsx", result);
    }

    [Fact]
    public void NormalizeWindowsPath_WithEmptyPath_ThrowsException()
    {
        // Arrange
        var input = "";

        // Act & Assert
        var exception = Assert.Throws<ExshellException>(() => PathConverter.NormalizeWindowsPath(input));
        Assert.Equal(ExitCodes.ArgumentError, exception.ExitCode);
    }

    [Fact]
    public void NormalizeWindowsPath_WithUnsupportedWslPath_ThrowsException()
    {
        // Arrange
        var input = "/home/user/file.txt";

        // Act & Assert
        var exception = Assert.Throws<ExshellException>(() => PathConverter.NormalizeWindowsPath(input));
        Assert.Equal(ExitCodes.ArgumentError, exception.ExitCode);
        Assert.Contains("Unsupported WSL path", exception.Message);
    }

    [Theory]
    [InlineData("/mnt/c/test/file.txt", "C")]
    [InlineData("/mnt/d/work/sample.xlsx", "D")]
    [InlineData("/mnt/e/data/", "E")]
    public void NormalizeWindowsPath_WithWslMountPath_ConvertsDriveCorrectly(string wslPath, string expectedDrive)
    {
        // Act
        var result = PathConverter.NormalizeWindowsPath(wslPath);

        // Assert
        Assert.StartsWith($"{expectedDrive}:", result);
    }

    [Fact]
    public void ToWslPath_WithWindowsPath_ReturnsWslPath()
    {
        // Arrange
        var input = @"C:\Temp\file.txt";

        // Act
        var result = PathConverter.ToWslPath(input);

        // Assert
        Assert.Equal("/mnt/c/Temp/file.txt", result);
    }

    [Theory]
    [InlineData(@"C:\Temp\file.txt", "/mnt/c/Temp/file.txt")]
    [InlineData(@"D:\work\sample.xlsx", "/mnt/d/work/sample.xlsx")]
    [InlineData(@"E:\data\test\", "/mnt/e/data/test")]
    public void ToWslPath_WithVariousPaths_ConvertsCorrectly(string windowsPath, string expectedWslPath)
    {
        // Act
        var result = PathConverter.ToWslPath(windowsPath);

        // Assert
        Assert.Equal(expectedWslPath, result);
    }
}