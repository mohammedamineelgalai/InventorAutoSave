// ============================================================================
// BuildInfo.cs - Auto version + build date helper
// ============================================================================
// Exposes runtime-resolved Version + BuildDate so window titles and footers
// auto-update at compile time. NEVER hardcode "v1.X.X - R YYYY-MM-DD" in XAML
// anymore; bind to {x:Static svc:BuildInfo.XXX} instead.
// ============================================================================

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace InventorAutoSave.Services;

public static class BuildInfo
{
    // --- Assembly metadata (read once at startup) ---
    public static string Version { get; } = ReadVersion();
    public static string BuildDate { get; } = ReadBuildDate();

    // --- Common formatted strings (used by XAML via x:Static) ---
    public const string Author    = "Mohammed Amine Elgalai";
    public const string Company   = "XNRGY CLIMATE SYSTEMS ULC";
    public const string TechStack = ".NET 8 | WPF | API COM Inventor";

    // Window titles + footers - "Inventor AutoSave - <Section> - v<X.Y.Z> - R <YYYY-MM-DD> - By <Author> - <Company>"
    public static string SettingsTitle { get; } = $"Inventor AutoSave - Configuration - v{Version} - R {BuildDate} - By {Author} - {Company}";
    public static string InfoTitle     { get; } = $"Inventor AutoSave - Info - v{Version} - R {BuildDate} - By {Author} - {Company}";

    // Short subtitle shown in window headers (e.g. "Smart Tools Amine v1.0 - XNRGY")
    public static string ShortSubtitle { get; } = $"Smart Tools Amine v{ShortVersion(Version)} - XNRGY";
    public static string InfoSubtitle  { get; } = $"Smart Tools Amine v{ShortVersion(Version)} - Sauvegarde automatique via API COM";

    // Tooltip + log line used by App.xaml.cs
    public static string AppToolTip   { get; } = $"Inventor AutoSave v{Version}";
    public static string StartupLog   { get; } = $"[+] InventorAutoSave v{Version} demarre (build {BuildDate})";

    // ───────────────────────────────────────────────────────────────────────

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
        // 1) Prefer the main module file timestamp (works in SingleFile publish)
        try
        {
            var path = Process.GetCurrentProcess().MainModule?.FileName;
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
                return File.GetLastWriteTime(path).ToString("yyyy-MM-dd");
        }
        catch { /* ignore */ }

        // 2) Fallback: assembly location (may be empty in SingleFile)
        try
        {
            var loc = Assembly.GetExecutingAssembly().Location;
            if (!string.IsNullOrEmpty(loc) && File.Exists(loc))
                return File.GetLastWriteTime(loc).ToString("yyyy-MM-dd");
        }
        catch { /* ignore */ }

        // 3) Last resort: today's date
        return DateTime.Now.ToString("yyyy-MM-dd");
    }

    // "1.0.0" -> "1.0"
    private static string ShortVersion(string fullVersion)
    {
        var parts = fullVersion.Split('.');
        return parts.Length >= 2 ? $"{parts[0]}.{parts[1]}" : fullVersion;
    }
}
