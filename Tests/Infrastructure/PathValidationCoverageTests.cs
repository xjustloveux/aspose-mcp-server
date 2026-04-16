using System.Text;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Meta-test that verifies all handler source files calling file I/O sinks also call a
///     SecurityHelper path-validation method. This acts as a regression guard — any new handler
///     that writes, reads, copies, or deletes files must include a SecurityHelper call, otherwise
///     this test fails with a clear list of non-compliant files.
/// </summary>
/// <remarks>
///     Approach: glob all Handlers/**/*.cs, scan source text for sink and validation patterns.
///     Handlers with no sinks pass unconditionally (in-memory-only handlers are fine).
///     Handlers with sinks but no validation call are reported as failures.
/// </remarks>
public class PathValidationCoverageTests
{
    /// <summary>
    ///     File I/O sink patterns that indicate a handler touches the filesystem with user-supplied paths.
    ///     Each pattern is a substring to check for in the handler source.
    /// </summary>
    private static readonly string[] SinkPatterns =
    [
        ".Save(",
        "File.Copy(",
        "File.Delete(",
        "File.WriteAllText(",
        "File.WriteAllBytes(",
        "File.ReadAllText(",
        "File.ReadAllBytes(",
        "File.OpenRead(",
        "new FileStream(",
        "File.Create(",
        "Directory.Delete("
    ];

    /// <summary>
    ///     SecurityHelper validation calls that satisfy the path-validation requirement.
    ///     A handler must contain at least one of these to be considered compliant.
    /// </summary>
    private static readonly string[] ValidationPatterns =
    [
        "SecurityHelper.ValidateFilePath(",
        "SecurityHelper.ResolveAndEnsureWithinAllowlist(",
        "SecurityHelper.ValidatePathWithinAllowedBases(",
        "SecurityHelper.ValidateUserPath(",
        "SecurityHelper.SafeRecursiveDelete("
    ];

    /// <summary>
    ///     Root directory containing handler source files. Resolved relative to the test assembly
    ///     location by walking up to find the solution root that contains "Handlers/".
    /// </summary>
    private static readonly string HandlersRoot = ResolveHandlersRoot();

    /// <summary>
    ///     Verifies that every handler source file which contains a file I/O sink also contains at
    ///     least one SecurityHelper path-validation call. Fails with an explicit list of non-compliant
    ///     files if any handler has sinks but no validation.
    /// </summary>
    [Fact]
    public void AllHandlersWithFileSinks_MustCallSecurityHelperValidation()
    {
        Assert.True(Directory.Exists(HandlersRoot),
            $"Handlers root directory not found: '{HandlersRoot}'. " +
            "The test cannot locate the handler source files.");

        var handlerFiles = Directory.GetFiles(HandlersRoot, "*.cs", SearchOption.AllDirectories)
            .OrderBy(f => f)
            .ToArray();

        Assert.NotEmpty(handlerFiles);

        var nonCompliant = new List<(string RelativePath, string[] SinksFound)>();

        foreach (var filePath in handlerFiles)
        {
            string source;
            try
            {
                source = File.ReadAllText(filePath, Encoding.UTF8);
            }
            catch (IOException ex)
            {
                throw new IOException($"Could not read handler source file '{filePath}': {ex.Message}", ex);
            }

            var sinksFound = SinkPatterns.Where(p => source.Contains(p, StringComparison.Ordinal)).ToArray();
            if (sinksFound.Length == 0)
                // No sinks — handler is in-memory-only, passes unconditionally.
                continue;

            var hasValidation = ValidationPatterns.Any(p => source.Contains(p, StringComparison.Ordinal));
            if (!hasValidation)
            {
                var relative = Path.GetRelativePath(HandlersRoot, filePath);
                nonCompliant.Add((relative, sinksFound));
            }
        }

        if (nonCompliant.Count == 0)
            return;

        var message = BuildFailureMessage(nonCompliant, handlerFiles.Length);
        Assert.Fail(message);
    }

    /// <summary>
    ///     Verifies that the Handlers root directory exists and contains at least one handler file,
    ///     so a misconfigured root produces a clear diagnostic rather than a false-green test.
    /// </summary>
    [Fact]
    public void HandlersRoot_Exists_And_ContainsHandlerFiles()
    {
        Assert.True(Directory.Exists(HandlersRoot),
            $"Handlers root directory not found: '{HandlersRoot}'.");

        var count = Directory.GetFiles(HandlersRoot, "*.cs", SearchOption.AllDirectories).Length;
        Assert.True(count > 0,
            $"No .cs handler files found under '{HandlersRoot}'. " +
            "Either the root path is wrong or handler files are missing.");
    }

    /// <summary>
    ///     Resolves the absolute path to the production Handlers/ directory by walking up from the
    ///     test assembly's bin location until a directory is found that contains BOTH a "Handlers"
    ///     subdirectory and the main project file "AsposeMcpServer.csproj". This distinguishes the
    ///     production handlers from the test Handlers/ directory that sits under Tests/.
    /// </summary>
    /// <returns>Absolute path to the production Handlers directory, or a fallback best-guess path.</returns>
    private static string ResolveHandlersRoot()
    {
        // Walk up from the test assembly output dir, looking for the solution/project root.
        // The production root has both "Handlers/" and "AsposeMcpServer.csproj" as children.
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var handlersCandidate = Path.Combine(dir.FullName, "Handlers");
            var csprojCandidate = Path.Combine(dir.FullName, "AsposeMcpServer.csproj");
            if (Directory.Exists(handlersCandidate) && File.Exists(csprojCandidate))
                return handlersCandidate;
            dir = dir.Parent;
        }

        // Fallback: best-guess relative to the known absolute project path.
        return "/home/node/aspose-mcp-server/Handlers";
    }

    /// <summary>
    ///     Builds a human-readable failure message listing all non-compliant handler files.
    /// </summary>
    /// <param name="nonCompliant">
    ///     List of tuples containing the relative file path and the sink patterns found in that file.
    /// </param>
    /// <param name="totalScanned">Total number of handler files scanned.</param>
    /// <returns>A formatted failure message string.</returns>
    private static string BuildFailureMessage(
        List<(string RelativePath, string[] SinksFound)> nonCompliant,
        int totalScanned)
    {
        var sb = new StringBuilder();
        sb.AppendLine(
            $"PATH VALIDATION COVERAGE FAILURE: {nonCompliant.Count} of {totalScanned} handler file(s) " +
            "contain file I/O sinks but NO SecurityHelper validation call.");
        sb.AppendLine();
        sb.AppendLine("Non-compliant handlers:");

        foreach (var (relativePath, sinksFound) in nonCompliant)
        {
            sb.AppendLine($"  - {relativePath}");
            sb.AppendLine($"      Sinks found: {string.Join(", ", sinksFound)}");
            sb.AppendLine($"      Required: at least one of [{string.Join(", ", ValidationPatterns)}]");
        }

        sb.AppendLine();
        sb.AppendLine("Fix: add SecurityHelper.ValidateFilePath() (or equivalent) before each file I/O sink.");
        sb.AppendLine("See: charter §5 red line — all file paths must pass SecurityHelper.ValidateFilePath().");
        return sb.ToString();
    }
}
