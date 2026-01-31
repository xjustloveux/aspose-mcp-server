using System.Runtime.InteropServices;
using Aspose.Pdf.Text;
using Aspose.Words;

namespace AsposeMcpServer.Core;

/// <summary>
///     Manages Aspose license loading and initialization.
///     Supports per-component independent license resolution to handle multiple .lic files.
/// </summary>
public static class LicenseManager
{
    /// <summary>
    ///     Sets Aspose licenses based on configuration.
    ///     Each component independently resolves its own license file with the following priority:
    ///     1. --license specified path (global override)
    ///     2. Component-specific .lic (e.g., Aspose.Words.lic)
    ///     3. Aspose.Total.lic (fallback)
    ///     4. Any .lic file found via directory scan (last resort)
    /// </summary>
    /// <param name="config">The server configuration.</param>
    public static void SetLicense(ServerConfig config)
    {
        var originalOut = Console.Out;
        try
        {
            Console.SetOut(TextWriter.Null);

            var globalLicensePath = ResolveGlobalLicensePath(config);
            var loadedLicenses = LoadLicenses(config, globalLicensePath);

            Console.SetOut(originalOut);
            ConfigureFontSubstitutions(config);
            LogLicenseResult(loadedLicenses);
        }
        catch (Exception ex)
        {
            Console.SetOut(originalOut);
            Console.Error.WriteLine($"[ERROR] Error loading Aspose license: {ex.Message}");
            Console.Error.WriteLine("[WARN] Running in evaluation mode.");
        }
    }

    /// <summary>
    ///     Configures font substitutions for Linux environments where common Windows fonts
    ///     (e.g., Arial, Times New Roman) are not available. Maps them to Liberation font equivalents.
    /// </summary>
    /// <param name="config">The server configuration.</param>
    private static void ConfigureFontSubstitutions(ServerConfig config)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            return;

        if (!config.EnablePdf)
            return;

        try
        {
            var fontDirs = new[]
            {
                "/usr/share/fonts",
                "/usr/share/fonts/truetype",
                "/usr/share/fonts/truetype/liberation"
            };

            foreach (var dir in fontDirs.Where(Directory.Exists))
                FontRepository.Sources.Add(new FolderFontSource(dir));

            var substitutions = new (string original, string replacement)[]
            {
                ("Arial", "Liberation Sans"),
                ("Times New Roman", "Liberation Serif"),
                ("Courier New", "Liberation Mono")
            };

            foreach (var (original, replacement) in substitutions)
                FontRepository.Substitutions.Add(new SimpleFontSubstitution(original, replacement));
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[WARN] Error configuring font substitutions: {ex.Message}");
        }
    }

    /// <summary>
    ///     Resolves the global license path from the --license command line argument or environment variable.
    /// </summary>
    /// <param name="config">The server configuration.</param>
    /// <returns>The resolved global license file path, or null if not specified or not found.</returns>
    private static string? ResolveGlobalLicensePath(ServerConfig config)
    {
        if (string.IsNullOrWhiteSpace(config.LicensePath)) return null;

        var baseDirectory = AppContext.BaseDirectory;
        var currentDirectory = Directory.GetCurrentDirectory();

        List<string> paths = [config.LicensePath];
        if (!Path.IsPathRooted(config.LicensePath))
        {
            paths.Add(Path.Combine(baseDirectory, config.LicensePath));
            paths.Add(Path.Combine(currentDirectory, config.LicensePath));
        }

        return paths.FirstOrDefault(File.Exists);
    }

    /// <summary>
    ///     Resolves the license file path for a specific component.
    ///     Search priority: global path → component .lic → Aspose.Total.lic → directory scan.
    /// </summary>
    /// <param name="componentLicFileName">The component-specific license filename (e.g., "Aspose.Words.lic").</param>
    /// <param name="globalLicensePath">The global license path from --license flag, or null.</param>
    /// <returns>The resolved license file path, or null if no license file is found.</returns>
    internal static string? ResolveComponentLicense(string componentLicFileName, string? globalLicensePath)
    {
        if (globalLicensePath != null)
            return globalLicensePath;

        var baseDirectory = AppContext.BaseDirectory;
        var currentDirectory = Directory.GetCurrentDirectory();

        var componentPaths = new[]
        {
            componentLicFileName,
            Path.Combine(baseDirectory, componentLicFileName),
            Path.Combine(currentDirectory, componentLicFileName)
        };
        var componentPath = componentPaths.FirstOrDefault(File.Exists);
        if (componentPath != null) return componentPath;

        var totalPaths = new[]
        {
            "Aspose.Total.lic",
            Path.Combine(baseDirectory, "Aspose.Total.lic"),
            Path.Combine(currentDirectory, "Aspose.Total.lic")
        };
        var totalPath = totalPaths.FirstOrDefault(File.Exists);
        if (totalPath != null) return totalPath;

        var searchDirs = new[] { baseDirectory, currentDirectory };
        foreach (var dir in searchDirs)
            try
            {
                var licFiles = Directory.GetFiles(dir, "*.lic", SearchOption.TopDirectoryOnly);
                if (licFiles.Length > 0) return licFiles[0];
            }
            catch
            {
                // Ignore directory access errors
            }

        return null;
    }

    /// <summary>
    ///     Loads licenses for enabled components. Each component independently resolves its own license file.
    /// </summary>
    /// <param name="config">The server configuration.</param>
    /// <param name="globalLicensePath">The global license path from --license flag, or null.</param>
    /// <returns>A list of successfully loaded license names.</returns>
    private static List<string> LoadLicenses(ServerConfig config, string? globalLicensePath)
    {
        List<string> loadedLicenses = [];

        var loaders = new (bool enabled, string name, string componentLic, Action<string> loader)[]
        {
            (config.EnableWord, "Words", "Aspose.Words.lic",
                p => new License().SetLicense(p)),
            (config.EnableExcel, "Cells", "Aspose.Cells.lic",
                p => new Aspose.Cells.License().SetLicense(p)),
            (config.EnablePowerPoint, "Slides", "Aspose.Slides.lic",
                p => new Aspose.Slides.License().SetLicense(p)),
            (config.EnablePdf, "Pdf", "Aspose.Pdf.lic",
                p => new Aspose.Pdf.License().SetLicense(p)),
            (config.EnableOcr, "OCR", "Aspose.OCR.lic",
                p => new Aspose.OCR.License().SetLicense(p))
        };

        foreach (var (enabled, name, componentLic, loader) in loaders)
        {
            if (!enabled) continue;
            var licensePath = ResolveComponentLicense(componentLic, globalLicensePath);
            if (licensePath == null) continue;
            try
            {
                loader(licensePath);
                loadedLicenses.Add(name);
            }
            catch
            {
                // Ignore license loading errors
            }
        }

        return loadedLicenses;
    }

    /// <summary>
    ///     Logs the license loading result.
    /// </summary>
    /// <param name="loadedLicenses">The list of successfully loaded licenses.</param>
    private static void LogLicenseResult(List<string> loadedLicenses)
    {
        if (loadedLicenses.Count > 0)
        {
            Console.Error.WriteLine($"[INFO] Aspose licenses loaded: {string.Join(", ", loadedLicenses)}");
        }
        else
        {
            Console.Error.WriteLine("[WARN] No Aspose licenses loaded. Running in evaluation mode.");
            Console.Error.WriteLine("[INFO] You can specify license file via:");
            Console.Error.WriteLine("[INFO]   - Environment variable: ASPOSE_LICENSE_PATH");
            Console.Error.WriteLine("[INFO]   - Command line: --license:path/to/license.lic");
        }
    }
}
