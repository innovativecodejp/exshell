namespace Exshell.Tests;

public class ExitCodesTests
{
    [Fact]
    public void ExitCodes_HaveExpectedValues()
    {
        // Assert - verify that exit codes match spec v0.3 requirements
        Assert.Equal(0, ExitCodes.Success);
        Assert.Equal(1, ExitCodes.ArgumentError);
        Assert.Equal(2, ExitCodes.SessionNotEstablished);
        Assert.Equal(3, ExitCodes.WorkbookNotFound);
        Assert.Equal(4, ExitCodes.WorksheetNotFound);
        Assert.Equal(5, ExitCodes.ShapeNotFound);
        Assert.Equal(6, ExitCodes.ExcelOperationFailed);
        Assert.Equal(7, ExitCodes.StandardInputFailed);
        Assert.Equal(8, ExitCodes.TempFileError);
        Assert.Equal(9, ExitCodes.DiffExecutionFailed);
        Assert.Equal(10, ExitCodes.UnexpectedError);
    }

    [Fact]
    public void ExitCodes_AreUnique()
    {
        // Arrange
        var exitCodes = new[]
        {
            ExitCodes.Success,
            ExitCodes.ArgumentError,
            ExitCodes.SessionNotEstablished,
            ExitCodes.WorkbookNotFound,
            ExitCodes.WorksheetNotFound,
            ExitCodes.ShapeNotFound,
            ExitCodes.ExcelOperationFailed,
            ExitCodes.StandardInputFailed,
            ExitCodes.TempFileError,
            ExitCodes.DiffExecutionFailed,
            ExitCodes.UnexpectedError
        };

        // Act
        var distinctCodes = exitCodes.Distinct().ToArray();

        // Assert
        Assert.Equal(exitCodes.Length, distinctCodes.Length);
    }

    [Theory]
    [InlineData(nameof(ExitCodes.Success), 0)]
    [InlineData(nameof(ExitCodes.ArgumentError), 1)]
    [InlineData(nameof(ExitCodes.SessionNotEstablished), 2)]
    [InlineData(nameof(ExitCodes.WorkbookNotFound), 3)]
    [InlineData(nameof(ExitCodes.WorksheetNotFound), 4)]
    [InlineData(nameof(ExitCodes.ShapeNotFound), 5)]
    [InlineData(nameof(ExitCodes.ExcelOperationFailed), 6)]
    public void ExitCodes_IndividualValues_AreCorrect(string codeName, int expectedValue)
    {
        // Use reflection to get the field value
        var field = typeof(ExitCodes).GetField(codeName);
        Assert.NotNull(field);
        
        var actualValue = field.GetValue(null);
        Assert.Equal(expectedValue, actualValue);
    }
}