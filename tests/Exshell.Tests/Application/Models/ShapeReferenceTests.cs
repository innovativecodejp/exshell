using Exshell.Application.Models;

namespace Exshell.Tests.Application.Models;

public class ShapeReferenceTests
{
    [Fact]
    public void Parse_WithSheetAndShapeName_ReturnsCorrectReference()
    {
        // Arrange
        var input = "Main:txtInput";
        var defaultSheet = "Default";

        // Act
        var result = ShapeReference.Parse(input, defaultSheet);

        // Assert
        Assert.Equal("Main", result.SheetName);
        Assert.Equal("txtInput", result.ShapeName);
    }

    [Fact]
    public void Parse_WithShapeNameOnly_UsesDefaultSheet()
    {
        // Arrange
        var input = "txtOutput";
        var defaultSheet = "Sheet1";

        // Act
        var result = ShapeReference.Parse(input, defaultSheet);

        // Assert
        Assert.Equal("Sheet1", result.SheetName);
        Assert.Equal("txtOutput", result.ShapeName);
    }

    [Fact]
    public void Parse_WithShapeNameOnlyAndNullDefault_UsesNullSheet()
    {
        // Arrange
        var input = "txtInput";
        string? defaultSheet = null;

        // Act
        var result = ShapeReference.Parse(input, defaultSheet);

        // Assert
        Assert.Null(result.SheetName);
        Assert.Equal("txtInput", result.ShapeName);
    }

    [Fact]
    public void Parse_WithEmptySheetName_ReturnsEmptySheet()
    {
        // Arrange
        var input = ":txtInput";
        var defaultSheet = "Default";

        // Act
        var result = ShapeReference.Parse(input, defaultSheet);

        // Assert
        Assert.Equal("", result.SheetName);
        Assert.Equal("txtInput", result.ShapeName);
    }

    [Fact]
    public void Parse_WithMultipleColons_UsesFirstColon()
    {
        // Arrange
        var input = "Sheet:Shape:Name";
        var defaultSheet = "Default";

        // Act
        var result = ShapeReference.Parse(input, defaultSheet);

        // Assert
        Assert.Equal("Sheet", result.SheetName);
        Assert.Equal("Shape:Name", result.ShapeName);
    }

    [Theory]
    [InlineData("Main:txtInput", "Main:txtInput")]
    [InlineData("txtOutput", "Sheet1:txtOutput")]
    public void ToString_ReturnsExpectedFormat(string input, string expected)
    {
        // Arrange
        var defaultSheet = "Sheet1";
        var shapeRef = ShapeReference.Parse(input, defaultSheet);

        // Act
        var result = shapeRef.ToString();

        // Assert
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ToString_WithNullSheetName_ReturnsShapeNameOnly()
    {
        // Arrange
        var shapeRef = ShapeReference.Parse("txtInput", null);

        // Act
        var result = shapeRef.ToString();

        // Assert
        Assert.Equal("txtInput", result);
    }

    [Theory]
    [InlineData("差分:左側", "差分", "左側")]
    [InlineData("メイン:入力テキスト", "メイン", "入力テキスト")]
    [InlineData("日本語シート名:日本語図形名", "日本語シート名", "日本語図形名")]
    public void Parse_WithJapaneseNames_HandlesCorrectly(string input, string expectedSheet, string expectedShape)
    {
        // Act
        var result = ShapeReference.Parse(input, "デフォルト");

        // Assert
        Assert.Equal(expectedSheet, result.SheetName);
        Assert.Equal(expectedShape, result.ShapeName);
    }

    [Theory]
    [InlineData("Sheet with spaces:Shape with spaces")]
    [InlineData("Sheet_with_underscores:Shape_with_underscores")]
    [InlineData("Sheet-with-hyphens:Shape-with-hyphens")]
    public void Parse_WithSpecialCharacters_HandlesCorrectly(string input)
    {
        // Act
        var result = ShapeReference.Parse(input, "Default");

        // Assert
        Assert.NotNull(result.SheetName);
        Assert.NotNull(result.ShapeName);
        Assert.Contains(":", input);
    }
}