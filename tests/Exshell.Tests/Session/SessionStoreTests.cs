using Exshell.Session;
using System.Text.Json;

namespace Exshell.Tests.Session;

public class SessionStoreTests : IDisposable
{
    private readonly string _tempDirectory;
    private readonly string _originalAppData;

    public SessionStoreTests()
    {
        // Create a temporary directory for testing
        _tempDirectory = Path.Combine(Path.GetTempPath(), "ExshellTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(_tempDirectory);

        // Override APPDATA environment variable for testing
        _originalAppData = Environment.GetEnvironmentVariable("APPDATA") ?? "";
        Environment.SetEnvironmentVariable("APPDATA", _tempDirectory);
    }

    public void Dispose()
    {
        // Restore original APPDATA
        Environment.SetEnvironmentVariable("APPDATA", _originalAppData);

        // Clean up temporary directory
        try
        {
            if (Directory.Exists(_tempDirectory))
                Directory.Delete(_tempDirectory, true);
        }
        catch
        {
            // Ignore cleanup errors in tests
        }
    }

    [Fact]
    public void Save_WithValidSession_CreatesSessionFile()
    {
        // Arrange
        var session = new SessionInfo
        {
            WorkbookPath = @"D:\work\sample.xlsx",
            DefaultSheetName = "Main"
        };

        // Act
        SessionStore.Save(session);

        // Assert
        var sessionPath = Path.Combine(_tempDirectory, "Exshell", "session.json");
        Assert.True(File.Exists(sessionPath));

        var json = File.ReadAllText(sessionPath);
        var loaded = JsonSerializer.Deserialize<SessionInfo>(json);
        Assert.NotNull(loaded);
        Assert.Equal(session.WorkbookPath, loaded.WorkbookPath);
        Assert.Equal(session.DefaultSheetName, loaded.DefaultSheetName);
    }

    [Fact]
    public void Load_WithExistingSession_ReturnsSession()
    {
        // Arrange
        var originalSession = new SessionInfo
        {
            WorkbookPath = @"D:\work\test.xlsx",
            DefaultSheetName = "Sheet1"
        };
        SessionStore.Save(originalSession);

        // Act
        var loadedSession = SessionStore.Load();

        // Assert
        Assert.NotNull(loadedSession);
        Assert.Equal(originalSession.WorkbookPath, loadedSession.WorkbookPath);
        Assert.Equal(originalSession.DefaultSheetName, loadedSession.DefaultSheetName);
    }

    [Fact]
    public void Load_WithNoSessionFile_ReturnsNull()
    {
        // Act
        var result = SessionStore.Load();

        // Assert
        Assert.Null(result);
    }

    [Fact]
    public void LoadOrThrow_WithExistingSession_ReturnsSession()
    {
        // Arrange
        var originalSession = new SessionInfo
        {
            WorkbookPath = @"D:\work\test.xlsx",
            DefaultSheetName = "Main"
        };
        SessionStore.Save(originalSession);

        // Act
        var loadedSession = SessionStore.LoadOrThrow();

        // Assert
        Assert.NotNull(loadedSession);
        Assert.Equal(originalSession.WorkbookPath, loadedSession.WorkbookPath);
        Assert.Equal(originalSession.DefaultSheetName, loadedSession.DefaultSheetName);
    }

    [Fact]
    public void LoadOrThrow_WithNoSession_ThrowsException()
    {
        // Arrange - ensure no session file exists
        var sessionPath = Path.Combine(_tempDirectory, "Exshell", "session.json");
        if (File.Exists(sessionPath))
        {
            File.Delete(sessionPath);
        }
        
        // Also clean up the entire Exshell directory if it exists
        var exshellDir = Path.Combine(_tempDirectory, "Exshell");
        if (Directory.Exists(exshellDir))
        {
            Directory.Delete(exshellDir, true);
        }


        // Act & Assert
        var exception = Assert.Throws<ExshellException>(() => SessionStore.LoadOrThrow());
        Assert.Equal(ExitCodes.SessionNotEstablished, exception.ExitCode);
        Assert.Contains("No active session", exception.Message);
    }

    [Fact]
    public void LoadOrThrow_WithEmptyWorkbookPath_ThrowsException()
    {
        // Arrange
        var invalidSession = new SessionInfo
        {
            WorkbookPath = "",
            DefaultSheetName = "Main"
        };
        SessionStore.Save(invalidSession);

        // Act & Assert
        var exception = Assert.Throws<ExshellException>(() => SessionStore.LoadOrThrow());
        Assert.Equal(ExitCodes.SessionNotEstablished, exception.ExitCode);
    }

    [Fact]
    public void Save_UpdatesTimestamp()
    {
        // Arrange
        var session = new SessionInfo
        {
            WorkbookPath = @"D:\work\sample.xlsx",
            DefaultSheetName = "Main"
        };
        var beforeSave = DateTime.Now;

        // Act
        SessionStore.Save(session);

        // Assert
        Assert.True(session.UpdatedAt >= beforeSave);
        Assert.True(session.UpdatedAt <= DateTime.Now);
    }
}