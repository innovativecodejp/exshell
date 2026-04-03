using System.Text.RegularExpressions;

namespace Exshell.Infrastructure;

/// <summary>
/// PowerShell / WSL 間のパス表現差を吸収する。
/// </summary>
public static partial class PathConverter
{
    public static string NormalizeWindowsPath(string inputPath)
    {
        if (string.IsNullOrWhiteSpace(inputPath))
        {
            throw new ExshellException("Path is required.", ExitCodes.ArgumentError);
        }

        var trimmed = TrimEnclosingQuotes(inputPath.Trim());
        var wslMatch = WslMountedPathPattern().Match(trimmed);

        if (wslMatch.Success)
        {
            var drive = wslMatch.Groups["drive"].Value.ToUpperInvariant();
            var rest = wslMatch.Groups["rest"].Value.Replace('/', '\\');
            return Path.GetFullPath($"{drive}:{rest}");
        }

        if (trimmed.StartsWith('/'))
        {
            throw new ExshellException(
                $"Unsupported WSL path: {trimmed}. Use a Windows path or /mnt/<drive>/... path.",
                ExitCodes.ArgumentError
            );
        }

        if (WindowsDrivePathPattern().IsMatch(trimmed))
        {
            trimmed = trimmed.Replace('/', '\\');
        }

        return Path.GetFullPath(trimmed);
    }

    /// <summary>
    /// Windows パスを WSL パスへ変換する。
    /// 例: C:\Temp\a.txt → /mnt/c/Temp/a.txt
    /// </summary>
    public static string ToWslPath(string windowsPath)
    {
        var normalized = NormalizeWindowsPath(windowsPath);

        if (normalized.Length >= 2 && normalized[1] == ':')
        {
            var drive = char.ToLowerInvariant(normalized[0]);
            var rest  = normalized[2..].Replace('\\', '/');
            return $"/mnt/{drive}{rest}";
        }

        return normalized.Replace('\\', '/');
    }

    private static string TrimEnclosingQuotes(string path)
    {
        if (path.Length >= 2 && path[0] == '"' && path[^1] == '"')
        {
            return path[1..^1];
        }

        return path;
    }

    [GeneratedRegex(@"^/mnt/(?<drive>[a-zA-Z])(?<rest>/.*)?$", RegexOptions.CultureInvariant)]
    private static partial Regex WslMountedPathPattern();

    [GeneratedRegex(@"^[a-zA-Z]:([\\/]|$)", RegexOptions.CultureInvariant)]
    private static partial Regex WindowsDrivePathPattern();
}
