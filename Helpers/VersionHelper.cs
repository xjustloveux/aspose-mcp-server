using System.Reflection;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Helper class to get application version from assembly info.
/// </summary>
public static class VersionHelper
{
    /// <summary>
    ///     Default version used when version cannot be determined.
    /// </summary>
    private const string DefaultVersion = "1.0.0";

    /// <summary>
    ///     Cached version string to avoid repeated assembly lookups.
    /// </summary>
    private static string? _version;

    /// <summary>
    ///     Gets the application version from assembly info.
    ///     Prefers InformationalVersion (preserves semantic version with prerelease tags),
    ///     falls back to AssemblyVersion, then to default "1.0.0".
    /// </summary>
    /// <returns>Version string in semantic version format (e.g., "1.2.3" or "1.2.3-beta").</returns>
    public static string GetVersion()
    {
        if (_version != null) return _version;

        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            var infoVersion = assembly
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                ?.InformationalVersion;

            if (!string.IsNullOrEmpty(infoVersion))
            {
                var plusIndex = infoVersion.IndexOf('+');
                _version = plusIndex > 0 ? infoVersion[..plusIndex] : infoVersion;
                return _version;
            }

            var version = assembly.GetName().Version;
            if (version != null)
            {
                _version = $"{version.Major}.{version.Minor}.{version.Build}";
                return _version;
            }
        }
        catch
        {
            // Ignore version retrieval errors, use default version
        }

        _version = DefaultVersion;
        return _version;
    }
}
