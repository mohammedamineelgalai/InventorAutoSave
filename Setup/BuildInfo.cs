// ============================================================================
// BuildInfo.cs (Setup) - Auto version + build date for the installer window
// ============================================================================

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace InventorAutoSave.Setup;

public static class BuildInfo
{
    public static string Version   { get; } = ReadVersion();
    public static string BuildDate { get; } = ReadBuildDate();

    public const string Author    = "Mohammed Amine Elgalai";
    public const string Company   = "XNRGY CLIMATE SYSTEMS ULC";

    // Title shown in the Window chrome (with em-dashes)
    public static string SetupTitle  { get; } = $"InventorAutoSave \u2014 Installation \u2014 v{Version} \u2014 R {BuildDate} \u2014 By {Author} \u2014 {Company}";
    // Footer shown above the Install/Cancel buttons
    public static string SetupFooter { get; } = $"InventorAutoSave \u2014 v{Version} \u2014 R {BuildDate} \u2014 By {Author} \u2014 {Company}";

    private static string ReadVersion()
    {
        try
        {
            var v = Assembly.GetExecutingAssembly().GetName().Version;
            if (v != null) return $"{v.Major}.{v.Minor}.{v.Build}";
        }
        catch { /* ignore */ }
        return "1.0.0";
    }

    private static string ReadBuildDate()
    {
        try
        {
            var path = Process.GetCurrentProcess().MainModule?.FileName;
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
                return File.GetLastWriteTime(path).ToString("yyyy-MM-dd");
        }
        catch { /* ignore */ }

        try
        {
            var loc = Assembly.GetExecutingAssembly().Location;
            if (!string.IsNullOrEmpty(loc) && File.Exists(loc))
                return File.GetLastWriteTime(loc).ToString("yyyy-MM-dd");
        }
        catch { /* ignore */ }

        return DateTime.Now.ToString("yyyy-MM-dd");
    }
}
