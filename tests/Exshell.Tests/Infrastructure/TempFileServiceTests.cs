using Exshell.Infrastructure;
using System.Text;

namespace Exshell.Tests.Infrastructure;

public class TempFileServiceTests
{
    [Fact]
    public void WriteUtf8Lf_WithSimpleText_CreatesFile()
    {
        // Arrange
        var text = "Hello, World!";
        string? tempFile = null;

        try
        {
            // Act
            tempFile = TempFileService.WriteUtf8Lf(text);

            // Assert
            Assert.NotNull(tempFile);
            Assert.True(File.Exists(tempFile));
            
            var content = File.ReadAllText(tempFile, Encoding.UTF8);
            Assert.Equal(text, content);
        }
        finally
        {
            // Cleanup
            TempFileService.TryDelete(tempFile);
        }
    }

    [Fact]
    public void WriteUtf8Lf_WithMultilineText_NormalizesToLf()
    {
        // Arrange
        var text = "Line 1\r\nLine 2\rLine 3\nLine 4";
        var expected = "Line 1\nLine 2\nLine 3\nLine 4";
        string? tempFile = null;

        try
        {
            // Act
            tempFile = TempFileService.WriteUtf8Lf(text);

            // Assert
            Assert.NotNull(tempFile);
            Assert.True(File.Exists(tempFile));
            
            var content = File.ReadAllText(tempFile, Encoding.UTF8);
            Assert.Equal(expected, content);
        }
        finally
        {
            // Cleanup
            TempFileService.TryDelete(tempFile);
        }
    }

    [Fact]
    public void WriteUtf8Lf_WithJapaneseText_PreservesCharacters()
    {
        // Arrange
        var text = "こんにちは\r\n世界！\nテスト";
        var expected = "こんにちは\n世界！\nテスト";
        string? tempFile = null;

        try
        {
            // Act
            tempFile = TempFileService.WriteUtf8Lf(text);

            // Assert
            Assert.NotNull(tempFile);
            Assert.True(File.Exists(tempFile));
            
            var content = File.ReadAllText(tempFile, Encoding.UTF8);
            Assert.Equal(expected, content);
        }
        finally
        {
            // Cleanup
            TempFileService.TryDelete(tempFile);
        }
    }

    [Fact]
    public void WriteUtf8Lf_WithEmptyText_CreatesEmptyFile()
    {
        // Arrange
        var text = "";
        string? tempFile = null;

        try
        {
            // Act
            tempFile = TempFileService.WriteUtf8Lf(text);

            // Assert
            Assert.NotNull(tempFile);
            Assert.True(File.Exists(tempFile));
            
            var content = File.ReadAllText(tempFile, Encoding.UTF8);
            Assert.Equal(text, content);
            
            var fileInfo = new FileInfo(tempFile);
            Assert.Equal(0, fileInfo.Length);
        }
        finally
        {
            // Cleanup
            TempFileService.TryDelete(tempFile);
        }
    }

    [Fact]
    public void WriteUtf8Lf_CreatesUtf8WithoutBom()
    {
        // Arrange
        var text = "UTF-8 BOM test";
        string? tempFile = null;

        try
        {
            // Act
            tempFile = TempFileService.WriteUtf8Lf(text);

            // Assert
            var bytes = File.ReadAllBytes(tempFile);
            
            // UTF-8 BOM is EF BB BF
            Assert.True(bytes.Length >= 3);
            Assert.False(bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF);
        }
        finally
        {
            // Cleanup
            TempFileService.TryDelete(tempFile);
        }
    }

    [Fact]
    public void TryDelete_WithExistingFile_DeletesFile()
    {
        // Arrange
        var tempFile = TempFileService.WriteUtf8Lf("test content");
        Assert.True(File.Exists(tempFile));

        // Act
        TempFileService.TryDelete(tempFile);

        // Assert
        Assert.False(File.Exists(tempFile));
    }

    [Fact]
    public void TryDelete_WithNullPath_DoesNotThrow()
    {
        // Act & Assert - should not throw
        TempFileService.TryDelete(null);
    }

    [Fact]
    public void TryDelete_WithNonExistentFile_DoesNotThrow()
    {
        // Arrange
        var nonExistentFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());

        // Act & Assert - should not throw
        TempFileService.TryDelete(nonExistentFile);
    }

    [Fact]
    public void WriteUtf8Lf_MultipleFiles_CreatesDifferentPaths()
    {
        // Arrange
        string? tempFile1 = null;
        string? tempFile2 = null;

        try
        {
            // Act
            tempFile1 = TempFileService.WriteUtf8Lf("Content 1");
            tempFile2 = TempFileService.WriteUtf8Lf("Content 2");

            // Assert
            Assert.NotNull(tempFile1);
            Assert.NotNull(tempFile2);
            Assert.NotEqual(tempFile1, tempFile2);
            
            Assert.True(File.Exists(tempFile1));
            Assert.True(File.Exists(tempFile2));
            
            Assert.Equal("Content 1", File.ReadAllText(tempFile1, Encoding.UTF8));
            Assert.Equal("Content 2", File.ReadAllText(tempFile2, Encoding.UTF8));
        }
        finally
        {
            // Cleanup
            TempFileService.TryDelete(tempFile1);
            TempFileService.TryDelete(tempFile2);
        }
    }
}