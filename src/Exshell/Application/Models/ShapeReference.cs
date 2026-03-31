namespace Exshell.Application.Models;

/// <summary>
/// コマンド引数から解決されたテキストボックス参照。
/// "SheetName:ShapeName" または "ShapeName" 形式を解析する。
/// </summary>
public sealed class ShapeReference
{
    public string? SheetName { get; }
    public string  ShapeName { get; }

    private ShapeReference(string? sheetName, string shapeName)
    {
        SheetName = sheetName;
        ShapeName = shapeName;
    }

    /// <summary>
    /// "Sheet:Shape" または "Shape" を解析する。
    /// シート名省略時は <paramref name="defaultSheet"/> を使用する。
    /// </summary>
    public static ShapeReference Parse(string input, string? defaultSheet)
    {
        var colon = input.IndexOf(':');
        if (colon >= 0)
            return new ShapeReference(input[..colon], input[(colon + 1)..]);

        return new ShapeReference(defaultSheet, input);
    }

    public override string ToString() =>
        SheetName != null ? $"{SheetName}:{ShapeName}" : ShapeName;
}
