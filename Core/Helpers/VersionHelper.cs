using System.Reflection;

namespace AsposeMcpServer.Core.Helpers;

/// <summary>
///     Helper class to get application version from assembly info
/// </summary>
public static class VersionHelper
{
    /// <summary>
    ///     Cached version string to avoid repeated assembly lookups
    /// </summary>
    private static string? _version;

    /// <summary>
    ///     Gets the application version from assembly info.
    ///     Falls back to "1.0.0" if version cannot be determined.
    /// </summary>
    /// <returns>Version string in format "major.minor.patch".</returns>
    public static string GetVersion()
    {
        if (_version != null) return _version;

        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            var version = assembly.GetName().Version;

            if (version != null)
            {
                var patch = version.Build >= 0 ? version.Build : version.Revision;
                _version = $"{version.Major}.{version.Minor}.{patch}";
                return _version;
            }
        }
        catch
        {
            // Ignore version retrieval errors, use default version
        }

        _version = "1.0.0";
        return _version;
    }
}
