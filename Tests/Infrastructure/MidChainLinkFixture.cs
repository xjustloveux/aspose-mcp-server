using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Creates a directory link whose creation does not require elevated privileges on Windows.
///     <see cref="Directory.CreateSymbolicLink" /> needs Developer Mode or administrator rights,
///     while an NTFS junction (<c>mklink /J</c>) does not, so a mid-chain redirection can be
///     exercised on an ordinary developer machine and on CI. On non-Windows platforms the helper
///     falls back to a normal directory symbolic link.
/// </summary>
public static class MidChainLinkFixture
{
    /// <summary>
    ///     Attempts to create a directory link at <paramref name="linkPath" /> that resolves to
    ///     <paramref name="targetPath" />.
    /// </summary>
    /// <param name="linkPath">Absolute path of the link entry to create. Must not already exist.</param>
    /// <param name="targetPath">Absolute path of the existing directory the link should point to.</param>
    /// <returns><c>true</c> when the link was created and resolves; otherwise <c>false</c>.</returns>
    public static bool TryCreateDirectoryLink(string linkPath, string targetPath)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            try
            {
                Directory.CreateSymbolicLink(linkPath, targetPath);
                return Directory.Exists(linkPath);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return false;
            }

        try
        {
            var psi = new ProcessStartInfo("cmd.exe")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            psi.ArgumentList.Add("/c");
            psi.ArgumentList.Add("mklink");
            psi.ArgumentList.Add("/J");
            psi.ArgumentList.Add(linkPath);
            psi.ArgumentList.Add(targetPath);

            using var process = Process.Start(psi);
            if (process == null) return false;
            process.WaitForExit(15000);
            return Directory.Exists(linkPath) && new DirectoryInfo(linkPath).LinkTarget != null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or Win32Exception)
        {
            return false;
        }
    }
}
