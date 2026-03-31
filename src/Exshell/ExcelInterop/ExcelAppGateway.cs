using System.Runtime.InteropServices;
using XL = Microsoft.Office.Interop.Excel;

namespace Exshell.ExcelInterop;

/// <summary>
/// Excel.Application の取得・生成を担当する。
/// Marshal.GetActiveObject は .NET 8 で廃止のため P/Invoke で代替する。
/// </summary>
public static class ExcelAppGateway
{
    [DllImport("ole32.dll")]
    private static extern int GetActiveObject(
        ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid pclsid);

    private const int S_OK = 0;

    /// <summary>
    /// 実行中の Excel.Application を返す。未起動なら null。
    /// </summary>
    public static XL.Application? TryGetRunningExcel()
    {
        if (CLSIDFromProgID("Excel.Application", out var clsid) != S_OK)
            return null;
        if (GetActiveObject(ref clsid, IntPtr.Zero, out var obj) != S_OK)
            return null;
        return obj as XL.Application;
    }

    /// <summary>
    /// 実行中の Excel を取得するか、なければ新規起動して返す。
    /// </summary>
    public static XL.Application GetOrCreateApplication()
    {
        var running = TryGetRunningExcel();
        if (running != null)
            return running;

        var app = new XL.Application();
        app.Visible = true;
        return app;
    }
}
