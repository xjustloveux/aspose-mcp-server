using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Pdf;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tools.Session;
using License = Aspose.Words.License;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Base class for all tests providing common functionality
/// </summary>
public abstract class TestBase : IDisposable
{
    /// <summary>
    ///     Aspose library types for evaluation mode checking
    /// </summary>
    public enum AsposeLibraryType
    {
        Slides,
        Words,
        Cells,
        Pdf
    }

    /// <summary>
    ///     Tracks which Aspose licenses were successfully loaded during initialization.
    ///     Used by IsEvaluationMode() to determine if a specific library is licensed.
    /// </summary>
    private static readonly HashSet<AsposeLibraryType> LoadedLicenses = new();

    protected readonly string TestDir;
    protected readonly List<string> TestFiles = new();

    /// <summary>
    ///     Shared session manager for session-based tests (lazy initialized)
    /// </summary>
    private DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Session tool for opening/closing sessions (lazy initialized)
    /// </summary>
    private DocumentSessionTool? _sessionTool;

    /// <summary>
    ///     Shared temp file manager for session-based tests (lazy initialized)
    /// </summary>
    private TempFileManager? _tempFileManager;

    static TestBase()
    {
        LoadAsposeLicenses();
    }

    protected TestBase()
    {
        TestDir = Path.Combine(Path.GetTempPath(), "AsposeMcpServerTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(TestDir);
    }

    /// <summary>
    ///     Gets the shared session config
    /// </summary>
    private static SessionConfig SessionConfig => new() { Enabled = true };

    /// <summary>
    ///     Gets the shared session manager (creates if not exists)
    /// </summary>
    protected DocumentSessionManager SessionManager =>
        _sessionManager ??= new DocumentSessionManager(SessionConfig);

    /// <summary>
    ///     Gets the shared temp file manager (creates if not exists)
    /// </summary>
    protected TempFileManager TempFileManager =>
        _tempFileManager ??= new TempFileManager(SessionConfig);

    /// <summary>
    ///     Gets the shared session tool (creates if not exists)
    /// </summary>
    protected DocumentSessionTool SessionTool =>
        _sessionTool ??= new DocumentSessionTool(SessionManager, TempFileManager, new StdioSessionIdentityAccessor());

    public virtual void Dispose()
    {
        _sessionManager?.Dispose();
        _tempFileManager?.Dispose();

        foreach (var file in TestFiles) DeleteFileWithRetry(file);

        DeleteDirectoryWithRetry(TestDir);
    }

    /// <summary>
    ///     Opens a document session and returns the session ID
    /// </summary>
    protected string OpenSession(string documentPath)
    {
        var result = SessionTool.Execute("open", documentPath);
        var json = JsonDocument.Parse(result);
        return json.RootElement.GetProperty("sessionId").GetString()!;
    }

    /// <summary>
    ///     Loads Aspose licenses if available and not skipped via environment variable.
    ///     The SKIP_ASPOSE_LICENSE environment variable is set by test.ps1 when -SkipLicense parameter is used.
    /// </summary>
    private static void LoadAsposeLicenses()
    {
        var skipLicense = Environment.GetEnvironmentVariable("SKIP_ASPOSE_LICENSE");
        if (string.Equals(skipLicense, "true", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(skipLicense, "1", StringComparison.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine("[INFO] SKIP_ASPOSE_LICENSE is set. Tests will run in evaluation mode.");
            return;
        }

        var baseDirectory = AppContext.BaseDirectory;
        var currentDirectory = Directory.GetCurrentDirectory();

        // Check environment variable (set by CI)
        var envLicensePath = Environment.GetEnvironmentVariable("ASPOSE_LICENSE_PATH");

        List<string> licenseFileNames = [];

        // Prioritize environment variable specified path
        if (!string.IsNullOrWhiteSpace(envLicensePath))
        {
            licenseFileNames.Add(envLicensePath);
            if (!Path.IsPathRooted(envLicensePath))
            {
                licenseFileNames.Add(Path.Combine(baseDirectory, envLicensePath));
                licenseFileNames.Add(Path.Combine(currentDirectory, envLicensePath));
            }
        }

        licenseFileNames.AddRange(
        [
            "Aspose.Total.lic",
            "Aspose.Words.lic",
            "Aspose.Cells.lic",
            "Aspose.Slides.lic",
            "Aspose.Pdf.lic",
            Path.Combine(baseDirectory, "Aspose.Total.lic"),
            Path.Combine(baseDirectory, "Aspose.Words.lic"),
            Path.Combine(baseDirectory, "Aspose.Cells.lic"),
            Path.Combine(baseDirectory, "Aspose.Slides.lic"),
            Path.Combine(baseDirectory, "Aspose.Pdf.lic"),
            Path.Combine(currentDirectory, "Aspose.Total.lic"),
            Path.Combine(currentDirectory, "Aspose.Words.lic"),
            Path.Combine(currentDirectory, "Aspose.Cells.lic"),
            Path.Combine(currentDirectory, "Aspose.Slides.lic"),
            Path.Combine(currentDirectory, "Aspose.Pdf.lic")
        ]);

        var searchDirectories = new[] { baseDirectory, currentDirectory };
        foreach (var dir in searchDirectories)
            try
            {
                var licFiles = Directory.GetFiles(dir, "*.lic", SearchOption.TopDirectoryOnly);
                licenseFileNames.AddRange(licFiles);
            }
            catch
            {
                // Ignore directory access errors
            }

        string? licensePath = null;
        foreach (var path in licenseFileNames.Distinct())
            if (File.Exists(path))
            {
                licensePath = path;
                break;
            }

        if (licensePath == null)
        {
            Console.Error.WriteLine("[WARN] No Aspose license file found. Tests will run in evaluation mode.");
            return;
        }

        try
        {
            var wordsLicense = new License();
            wordsLicense.SetLicense(licensePath);
            LoadedLicenses.Add(AsposeLibraryType.Words);
        }
        catch
        {
            // License file might not contain Words license
        }

        try
        {
            var cellsLicense = new Aspose.Cells.License();
            cellsLicense.SetLicense(licensePath);
            LoadedLicenses.Add(AsposeLibraryType.Cells);
        }
        catch
        {
            // License file might not contain Cells license
        }

        try
        {
            var slidesLicense = new Aspose.Slides.License();
            slidesLicense.SetLicense(licensePath);
            LoadedLicenses.Add(AsposeLibraryType.Slides);
        }
        catch
        {
            // License file might not contain Slides license
        }

        try
        {
            var pdfLicense = new Aspose.Pdf.License();
            pdfLicense.SetLicense(licensePath);
            LoadedLicenses.Add(AsposeLibraryType.Pdf);
        }
        catch
        {
            // License file might not contain Pdf license
        }

        if (LoadedLicenses.Count > 0)
        {
            Console.Error.WriteLine($"[INFO] Aspose licenses loaded successfully from: {licensePath}");
            Console.Error.WriteLine($"[INFO] Loaded licenses: {string.Join(", ", LoadedLicenses)}");
        }
        else
        {
            Console.Error.WriteLine($"[WARN] License file found but no valid licenses loaded: {licensePath}");
        }
    }

    /// <summary>
    ///     Deletes a file with retry mechanism to handle locked files
    /// </summary>
    private static void DeleteFileWithRetry(string filePath, int maxRetries = 3, int delayMs = 100)
    {
        if (!File.Exists(filePath))
            return;

        for (var attempt = 0; attempt < maxRetries; attempt++)
            try
            {
                // Force garbage collection to release file handles
                if (attempt > 0)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Thread.Sleep(delayMs * (attempt + 1));
                }

                File.Delete(filePath);
                return; // Success
            }
            catch (IOException)
            {
                // File is locked, retry
                if (attempt == maxRetries - 1)
                    // Last attempt failed, try to remove read-only attribute and retry
                    try
                    {
                        var fileInfo = new FileInfo(filePath);
                        if (fileInfo.Exists)
                        {
                            fileInfo.IsReadOnly = false;
                            File.Delete(filePath);
                            return;
                        }
                    }
                    catch
                    {
                        // Ignore final cleanup errors
                    }
            }
            catch (UnauthorizedAccessException)
            {
                // Permission denied, try to remove read-only attribute
                try
                {
                    var fileInfo = new FileInfo(filePath);
                    if (fileInfo.Exists)
                    {
                        fileInfo.IsReadOnly = false;
                        File.Delete(filePath);
                        return;
                    }
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
            catch
            {
                // Other errors, ignore after retries
                if (attempt == maxRetries - 1)
                    return;
            }
    }

    /// <summary>
    ///     Deletes a directory with retry mechanism to handle locked files
    /// </summary>
    private static void DeleteDirectoryWithRetry(string directoryPath, int maxRetries = 5, int delayMs = 200)
    {
        if (!Directory.Exists(directoryPath))
            return;

        for (var attempt = 0; attempt < maxRetries; attempt++)
            try
            {
                // Force garbage collection to release file handles
                if (attempt > 0)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect(); // Second collection to ensure finalizers completed
                    Thread.Sleep(delayMs * (attempt + 1));
                }

                // Try to delete all files first, then directory
                var files = Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories);
                foreach (var file in files) DeleteFileWithRetry(file, 2, 50);

                // Try to delete all subdirectories first
                var subDirs = Directory.GetDirectories(directoryPath, "*", SearchOption.AllDirectories);
                foreach (var subDir in subDirs.Reverse()) // Delete from deepest first
                    try
                    {
                        Directory.Delete(subDir, false);
                    }
                    catch
                    {
                        // Ignore individual subdirectory errors
                    }

                // Finally delete the main directory
                Directory.Delete(directoryPath, false);
                return; // Success
            }
            catch (IOException)
            {
                // Directory or files are locked, retry
                if (attempt == maxRetries - 1)
                    // Last attempt: try to remove read-only attributes
                    try
                    {
                        var dirInfo = new DirectoryInfo(directoryPath);
                        if (dirInfo.Exists)
                        {
                            RemoveReadOnlyAttributes(dirInfo);
                            Directory.Delete(directoryPath, true);
                            return;
                        }
                    }
                    catch
                    {
                        // Ignore final cleanup errors - directory will be cleaned up later
                    }
            }
            catch (UnauthorizedAccessException)
            {
                // Permission denied, try to remove read-only attributes
                try
                {
                    var dirInfo = new DirectoryInfo(directoryPath);
                    if (dirInfo.Exists)
                    {
                        RemoveReadOnlyAttributes(dirInfo);
                        Directory.Delete(directoryPath, true);
                        return;
                    }
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
            catch
            {
                // Other errors, ignore after retries
                if (attempt == maxRetries - 1)
                    return;
            }
    }

    /// <summary>
    ///     Removes read-only attributes from directory and files recursively
    /// </summary>
    private static void RemoveReadOnlyAttributes(DirectoryInfo dirInfo)
    {
        try
        {
            dirInfo.Attributes &= ~FileAttributes.ReadOnly;

            foreach (var file in dirInfo.GetFiles())
                try
                {
                    file.Attributes &= ~FileAttributes.ReadOnly;
                }
                catch
                {
                    // Ignore individual file errors
                }

            foreach (var subDir in dirInfo.GetDirectories()) RemoveReadOnlyAttributes(subDir);
        }
        catch
        {
            // Ignore errors
        }
    }

    /// <summary>
    ///     Checks if Aspose libraries are running in evaluation mode.
    ///     Evaluation mode may add watermarks and limit certain operations.
    ///     This allows tests to adapt behavior when running without a license.
    /// </summary>
    /// <param name="libraryType">The Aspose library type to check.</param>
    /// <returns>True if running in evaluation mode, false if licensed.</returns>
    protected static bool IsEvaluationMode(AsposeLibraryType libraryType = AsposeLibraryType.Slides)
    {
        // First check environment variable (set by test.ps1 -SkipLicense)
        var skipLicense = Environment.GetEnvironmentVariable("SKIP_ASPOSE_LICENSE")
                          ?? Environment.GetEnvironmentVariable("SKIP_LICENSE");
        if (string.Equals(skipLicense, "true", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(skipLicense, "1", StringComparison.OrdinalIgnoreCase))
            return true;

        // Use direct API checks for libraries that support it
        // This correctly handles expired licenses (SetLicense succeeded but license expired)
        try
        {
            return libraryType switch
            {
                // Aspose.Words: No direct IsLicensed API, use HashSet tracking
                AsposeLibraryType.Words => !LoadedLicenses.Contains(AsposeLibraryType.Words),
                // Aspose.Cells: Workbook.IsLicensed is an instance property
                AsposeLibraryType.Cells => !new Workbook().IsLicensed,
                // Aspose.Pdf: Document.IsLicensed is a STATIC property
                AsposeLibraryType.Pdf => !Document.IsLicensed,
                // Aspose.Slides: License.IsLicensed() is a method on License class
                AsposeLibraryType.Slides => !new Aspose.Slides.License().IsLicensed(),
                _ => !new Aspose.Slides.License().IsLicensed()
            };
        }
        catch
        {
            // If API check fails, fall back to HashSet tracking
            return !LoadedLicenses.Contains(libraryType);
        }
    }

    /// <summary>
    ///     Skips the current test if running in evaluation mode.
    ///     Use this with [SkippableFact] or [SkippableTheory] attribute.
    ///     This properly marks the test as "Skipped" instead of silently passing.
    /// </summary>
    /// <param name="libraryType">The Aspose library type to check.</param>
    /// <param name="reason">The reason for skipping (shown in test output).</param>
    protected static void SkipInEvaluationMode(
        AsposeLibraryType libraryType = AsposeLibraryType.Slides,
        string reason = "Exceeds evaluation mode limits")
    {
        Skip.If(IsEvaluationMode(libraryType), reason);
    }

    /// <summary>
    ///     Creates a test file path
    /// </summary>
    protected string CreateTestFilePath(string fileName)
    {
        var filePath = Path.Combine(TestDir, fileName);
        TestFiles.Add(filePath);
        return filePath;
    }

    /// <summary>
    ///     Creates a JsonObject with common parameters
    /// </summary>
    protected static JsonObject CreateArguments(string operation, string path, string? outputPath = null)
    {
        var args = new JsonObject
        {
            ["operation"] = operation,
            ["path"] = path
        };

        if (!string.IsNullOrEmpty(outputPath)) args["outputPath"] = outputPath;

        return args;
    }
}
