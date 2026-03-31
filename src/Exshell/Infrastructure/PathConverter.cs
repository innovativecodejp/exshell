namespace Exshell.Infrastructure;

/// <summary>
/// Windows パスを WSL パスへ変換する。
/// 例: C:\Temp\a.txt → /mnt/c/Temp/a.txt
/// </summary>
public static class PathConverter
{
    public static string ToWslPath(string windowsPath)
    {
        if (windowsPath.Length >= 2 && windowsPath[1] == ':')
        {
            var drive = char.ToLower(windowsPath[0]);
            var rest  = windowsPath[2..].Replace('\\', '/');
            return $"/mnt/{drive}{rest}";
        }
        return windowsPath.Replace('\\', '/');
    }
}
