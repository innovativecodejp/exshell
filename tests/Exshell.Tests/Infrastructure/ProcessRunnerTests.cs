using Exshell.Infrastructure;
using System.Diagnostics;

namespace Exshell.Tests.Infrastructure;

public class ProcessRunnerTests
{
    private static bool IsWslAvailable()
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "wsl",
                Arguments = "--version",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var proc = Process.Start(psi);
            proc?.WaitForExit(5000);
            return proc?.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }

    [Fact]
    public void RunWsl_WithEchoCommand_ReturnsSuccessCode()
    {
        // Skip test if WSL is not available
        if (!IsWslAvailable())
        {
            Assert.True(true, "WSL is not available - test skipped");
            return;
        }

        // Arrange
        var command = "echo 'test'";

        // Act
        var exitCode = ProcessRunner.RunWsl(command);

        // Assert
        Assert.Equal(0, exitCode);
    }

    [Fact]
    public void RunWsl_WithInvalidCommand_ReturnsNonZeroCode()
    {
        // Skip test if WSL is not available
        if (!IsWslAvailable())
        {
            Assert.True(true, "WSL is not available - test skipped");
            return;
        }

        // Arrange
        var command = "nonexistentcommand12345";

        // Act
        var exitCode = ProcessRunner.RunWsl(command);

        // Assert
        Assert.NotEqual(0, exitCode);
    }

    [Fact]
    public void RunWsl_WithDiffCommand_ReturnsCorrectCode()
    {
        // Skip test if WSL is not available
        if (!IsWslAvailable())
        {
            Assert.True(true, "WSL is not available - test skipped");
            return;
        }

        // Arrange - use a simple diff command that should succeed
        var command = "echo 'same' | diff - <(echo 'same')";

        // Act
        var exitCode = ProcessRunner.RunWsl(command);

        // Assert
        Assert.Equal(0, exitCode); // No differences should return 0
    }

    [Fact]
    public void RunWsl_WithNullOrEmptyCommand_ThrowsException()
    {
        // Act & Assert
        Assert.Throws<ExshellException>(() => ProcessRunner.RunWsl(""));
    }

    // Note: More comprehensive tests would require:
    // 1. Mocking the Process.Start method
    // 2. Testing stdout/stderr capture
    // 3. Testing UTF-8 encoding
    // 4. Testing timeout scenarios
    // 
    // These would require additional mocking frameworks like Moq
    // and are beyond the scope of basic unit tests without external dependencies.
}