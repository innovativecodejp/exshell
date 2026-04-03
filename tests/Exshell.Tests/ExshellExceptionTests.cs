namespace Exshell.Tests;

public class ExshellExceptionTests
{
    [Fact]
    public void Constructor_WithMessageAndExitCode_SetsProperties()
    {
        // Arrange
        var message = "Test error message";
        var exitCode = ExitCodes.ArgumentError;

        // Act
        var exception = new ExshellException(message, exitCode);

        // Assert
        Assert.Equal(message, exception.Message);
        Assert.Equal(exitCode, exception.ExitCode);
    }

    [Theory]
    [InlineData("Session not found", ExitCodes.SessionNotEstablished)]
    [InlineData("Workbook not found", ExitCodes.WorkbookNotFound)]
    [InlineData("Shape not found", ExitCodes.ShapeNotFound)]
    [InlineData("Excel operation failed", ExitCodes.ExcelOperationFailed)]
    public void Constructor_WithVariousMessages_SetsCorrectly(string message, int exitCode)
    {
        // Act
        var exception = new ExshellException(message, exitCode);

        // Assert
        Assert.Equal(message, exception.Message);
        Assert.Equal(exitCode, exception.ExitCode);
    }

    [Fact]
    public void ExshellException_InheritsFromException()
    {
        // Arrange
        var exception = new ExshellException("Test", ExitCodes.UnexpectedError);

        // Assert
        Assert.IsAssignableFrom<Exception>(exception);
    }

    [Fact]
    public void ExshellException_CanBeThrownAndCaught()
    {
        // Arrange
        var message = "Test exception";
        var exitCode = ExitCodes.WorkbookNotFound;

        // Act & Assert
        var exception = Assert.Throws<ExshellException>((Action)(() =>
        {
            throw new ExshellException(message, exitCode);
        }));

        Assert.Equal(message, exception.Message);
        Assert.Equal(exitCode, exception.ExitCode);
    }

    [Fact]
    public void ExshellException_CanBeCaughtAsBaseException()
    {
        // Arrange
        var message = "Base exception test";
        var exitCode = ExitCodes.TempFileError;

        // Act & Assert
        var exception = Assert.Throws<Exception>((Action)(() =>
        {
            throw new ExshellException(message, exitCode);
        }));

        Assert.IsType<ExshellException>(exception);
        var exshellException = (ExshellException)exception;
        Assert.Equal(message, exshellException.Message);
        Assert.Equal(exitCode, exshellException.ExitCode);
    }
}