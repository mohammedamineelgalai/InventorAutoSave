// ============================================================================
// AssemblyInfo.cs — Manual assembly attributes
// ============================================================================
// This file replaces the auto-generated obj\...\AssemblyInfo.cs.
// Reason: The Roslyn C# language server generates assembly attributes
// in-memory AND reads the physical file in obj\, causing CS0579
// "Duplicate attribute" errors in VS Code. By setting
// <GenerateAssemblyInfo>false</GenerateAssemblyInfo> in the .csproj
// and providing attributes here manually, we eliminate the duplication.
// ============================================================================

using System.Reflection;
using System.Runtime.Versioning;

// --- Assembly metadata (mirrors .csproj properties) ---
[assembly: AssemblyTitle("InventorAutoSave")]
[assembly: AssemblyProduct("InventorAutoSave")]
[assembly: AssemblyCompany("XNRGY Climate Systems ULC")]
[assembly: AssemblyDescription("Inventor AutoSave - Smart Tools Amine v1.0 - Sauvegarde automatique via API COM Inventor (4-phases anti-segment-corruption)")]
[assembly: AssemblyCopyright("© 2026 Mohammed Amine Elgalai - XNRGY Climate Systems ULC")]
[assembly: AssemblyConfiguration("Release")]

// --- Versioning ---
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0.0")]

// --- Target framework & platform ---
[assembly: TargetFramework(".NETCoreApp,Version=v8.0", FrameworkDisplayName = ".NET 8.0")]
[assembly: TargetPlatform("Windows7.0")]
[assembly: SupportedOSPlatform("Windows7.0")]
