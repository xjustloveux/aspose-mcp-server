using System.Runtime.InteropServices;
using Aspose.Pdf.Text;
using Aspose.Words;

namespace AsposeMcpServer.Core;

/// <summary>
///     Manages Aspose license loading and initialization.
/// </summary>
public static class LicenseManager
{
    /// <summary>
    ///     Sets Aspose licenses based on configuration.
    ///     Searches for license files in multiple locations and loads licenses for enabled components.
    /// </summary>
    /// <param name="config">The server configuration.</param>
    public static void SetLicense(ServerConfig config)
    {
        var originalOut = Console.Out;
        try
        {
            Console.SetOut(TextWriter.Null);

            var licenseFileNames = BuildLicenseSearchPaths(config);
            var licensePath = licenseFileNames.FirstOrDefault(File.Exists);
            var loadedLicenses = LoadLicenses(config, licensePath);

            Console.SetOut(originalOut);
            ConfigureFontSubstitutions(config);
            LogLicenseResult(licensePath, loadedLicenses, licenseFileNames);
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

            foreach (var dir in fontDirs)
                if (Directory.Exists(dir))
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
    ///     Builds a list of potential license file paths to search.
    /// </summary>
    /// <param name="config">The server configuration.</param>
    /// <returns>A list of license file paths to search.</returns>
    private static List<string> BuildLicenseSearchPaths(ServerConfig config)
    {
        var baseDirectory = AppContext.BaseDirectory;
        var currentDirectory = Directory.GetCurrentDirectory();
        List<string> paths = [];

        AddConfiguredLicensePath(paths, config.LicensePath, baseDirectory, currentDirectory);
        AddComponentLicensePaths(paths, config, baseDirectory, currentDirectory);
        AddTotalLicensePath(paths, baseDirectory, currentDirectory);
        AddDiscoveredLicenseFiles(paths, baseDirectory, currentDirectory);

        return paths;
    }

    /// <summary>
    ///     Adds the configured license path if specified.
    /// </summary>
    /// <param name="paths">The list of paths to add to.</param>
    /// <param name="licensePath">The configured license path.</param>
    /// <param name="baseDirectory">The base directory for the application.</param>
    /// <param name="currentDirectory">The current working directory.</param>
    private static void AddConfiguredLicensePath(List<string> paths, string? licensePath,
        string baseDirectory, string currentDirectory)
    {
        if (string.IsNullOrWhiteSpace(licensePath)) return;

        paths.Add(licensePath);
        if (Path.IsPathRooted(licensePath)) return;

        paths.Add(Path.Combine(baseDirectory, licensePath));
        paths.Add(Path.Combine(currentDirectory, licensePath));
    }

    /// <summary>
    ///     Adds component-specific license paths based on enabled features.
    /// </summary>
    /// <param name="paths">The list of paths to add to.</param>
    /// <param name="config">The server configuration.</param>
    /// <param name="baseDirectory">The base directory for the application.</param>
    /// <param name="currentDirectory">The current working directory.</param>
    private static void AddComponentLicensePaths(List<string> paths, ServerConfig config,
        string baseDirectory, string currentDirectory)
    {
        var components = new (bool enabled, string fileName)[]
        {
            (config.EnableWord, "Aspose.Words.lic"),
            (config.EnableExcel, "Aspose.Cells.lic"),
            (config.EnablePowerPoint, "Aspose.Slides.lic"),
            (config.EnablePdf, "Aspose.Pdf.lic")
        };

        foreach (var (enabled, fileName) in components)
        {
            if (!enabled) continue;
            paths.Add(fileName);
            paths.Add(Path.Combine(baseDirectory, fileName));
            paths.Add(Path.Combine(currentDirectory, fileName));
        }
    }

    /// <summary>
    ///     Adds the Aspose.Total license path.
    /// </summary>
    /// <param name="paths">The list of paths to add to.</param>
    /// <param name="baseDirectory">The base directory for the application.</param>
    /// <param name="currentDirectory">The current working directory.</param>
    private static void AddTotalLicensePath(List<string> paths, string baseDirectory, string currentDirectory)
    {
        paths.Add("Aspose.Total.lic");
        paths.Add(Path.Combine(baseDirectory, "Aspose.Total.lic"));
        paths.Add(Path.Combine(currentDirectory, "Aspose.Total.lic"));
    }

    /// <summary>
    ///     Discovers and adds any .lic files in the search directories.
    /// </summary>
    /// <param name="paths">The list of paths to add to.</param>
    /// <param name="baseDirectory">The base directory for the application.</param>
    /// <param name="currentDirectory">The current working directory.</param>
    private static void AddDiscoveredLicenseFiles(List<string> paths, string baseDirectory, string currentDirectory)
    {
        var searchDirectories = new[] { baseDirectory, currentDirectory };
        foreach (var dir in searchDirectories)
            try
            {
                var licFiles = Directory.GetFiles(dir, "*.lic", SearchOption.TopDirectoryOnly);
                foreach (var licFile in licFiles)
                {
                    var fileName = Path.GetFileName(licFile);
                    if (!paths.Contains(licFile) && !paths.Contains(fileName))
                        paths.Add(licFile);
                }
            }
            catch
            {
                // Ignore directory access errors
            }
    }

    /// <summary>
    ///     Loads licenses for enabled components.
    /// </summary>
    /// <param name="config">The server configuration.</param>
    /// <param name="licensePath">The path to the license file.</param>
    /// <returns>A list of successfully loaded license names.</returns>
    private static List<string> LoadLicenses(ServerConfig config, string? licensePath)
    {
        List<string> loadedLicenses = [];
        if (licensePath == null) return loadedLicenses;

        var loaders = new (bool enabled, string name, Action loader)[]
        {
            (config.EnableWord, "Words", () => new License().SetLicense(licensePath)),
            (config.EnableExcel, "Cells", () => new Aspose.Cells.License().SetLicense(licensePath)),
            (config.EnablePowerPoint, "Slides", () => new Aspose.Slides.License().SetLicense(licensePath)),
            (config.EnablePdf, "Pdf", () => new Aspose.Pdf.License().SetLicense(licensePath))
        };

        foreach (var (enabled, name, loader) in loaders)
        {
            if (!enabled) continue;
            try
            {
                loader();
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
    /// <param name="licensePath">The license file path that was used.</param>
    /// <param name="loadedLicenses">The list of successfully loaded licenses.</param>
    /// <param name="searchedPaths">The list of paths that were searched.</param>
    private static void LogLicenseResult(string? licensePath, List<string> loadedLicenses,
        List<string> searchedPaths)
    {
        if (licensePath != null)
            LogLoadedLicenses(licensePath, loadedLicenses);
        else
            LogNoLicenseFound(searchedPaths);
    }

    /// <summary>
    ///     Logs successfully loaded licenses.
    /// </summary>
    /// <param name="licensePath">The license file path.</param>
    /// <param name="loadedLicenses">The list of loaded license names.</param>
    private static void LogLoadedLicenses(string licensePath, List<string> loadedLicenses)
    {
        if (loadedLicenses.Count > 0)
        {
            Console.Error.WriteLine($"[INFO] Aspose licenses loaded successfully from: {licensePath}");
            Console.Error.WriteLine($"[INFO] Loaded licenses: {string.Join(", ", loadedLicenses)}");
        }
        else
        {
            Console.Error.WriteLine($"[WARN] License file found but no valid licenses loaded: {licensePath}");
            Console.Error.WriteLine("[WARN] Running in evaluation mode.");
        }
    }

    /// <summary>
    ///     Logs when no license file is found.
    /// </summary>
    /// <param name="searchedPaths">The list of paths that were searched.</param>
    private static void LogNoLicenseFound(List<string> searchedPaths)
    {
        Console.Error.WriteLine("[WARN] No Aspose license file found. Searched locations:");
        var pathsToShow = searchedPaths.Distinct().Take(10);
        foreach (var path in pathsToShow)
            Console.Error.WriteLine($"[WARN]   - {Path.GetFullPath(path)}");

        if (searchedPaths.Count > 10)
            Console.Error.WriteLine($"[WARN]   ... and {searchedPaths.Count - 10} more locations");

        Console.Error.WriteLine("[WARN] Running in evaluation mode.");
        Console.Error.WriteLine("[INFO] You can specify license file via:");
        Console.Error.WriteLine("[INFO]   - Environment variable: ASPOSE_LICENSE_PATH");
        Console.Error.WriteLine("[INFO]   - Command line: --license:path/to/license.lic");
    }
}
