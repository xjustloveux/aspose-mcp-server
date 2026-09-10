using System.Text.RegularExpressions;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Keeps CODE_QUALITY_EXCEPTIONS.md honest (DOC-QUALITY-01). Every exception in that
///     document points at a source file, and the document used to pin each one to a line number
///     captured months earlier. Nothing verified either, so entries survived the file being split
///     away and the line numbers drifted silently. These tests fail when a referenced file no
///     longer exists and when a line-number column is reintroduced.
/// </summary>
public class CodeQualityExceptionsDocTests
{
    /// <summary>Matches a file path written as inline code, e.g. <c>`Core/Session/X.cs`</c>.</summary>
    private static readonly Regex ReferencedFile = new(
        @"`([A-Za-z0-9_./-]+\.cs)`", RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    /// <summary>Matches the table header that carried the line-number column.</summary>
    private static readonly Regex LineNumberColumn = new(
        @"^\|\s*檔案\s*\|\s*行號\s*\|", RegexOptions.Compiled | RegexOptions.Multiline,
        TimeSpan.FromSeconds(5));

    /// <summary>
    ///     Walks up from the test binaries to the repository root so the scan works from any
    ///     working directory the runner chooses.
    /// </summary>
    /// <returns>The directory holding the exceptions document.</returns>
    private static DirectoryInfo RepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "CODE_QUALITY_EXCEPTIONS.md")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        return dir;
    }

    [Fact]
    public void EveryReferencedFile_ShouldStillExist()
    {
        var root = RepositoryRoot();
        var document = File.ReadAllText(Path.Combine(root.FullName, "CODE_QUALITY_EXCEPTIONS.md"));

        var referenced = ReferencedFile.Matches(document)
            .Select(match => match.Groups[1].Value)
            .Distinct(StringComparer.Ordinal)
            .ToList();

        // An empty reference list has nothing missing from it, so this has to know the document
        // still references files at all (R9-T01).
        Assert.True(referenced.Count > 0,
            "CODE_QUALITY_EXCEPTIONS.md references no files at all; this guard is reading the "
            + "wrong document, or its table format has changed.");

        var missing = referenced
            .Where(path => !File.Exists(Path.Combine(root.FullName, path)))
            .ToList();

        Assert.True(missing.Count == 0,
            "CODE_QUALITY_EXCEPTIONS.md references files that no longer exist: "
            + string.Join(", ", missing)
            + ". Remove the entry or point it at the file the code moved to.");
    }

    [Fact]
    public void Document_ShouldNotReintroduceLineNumbers()
    {
        var root = RepositoryRoot();
        var document = File.ReadAllText(Path.Combine(root.FullName, "CODE_QUALITY_EXCEPTIONS.md"));

        Assert.False(LineNumberColumn.IsMatch(document),
            "A line-number column was added back to CODE_QUALITY_EXCEPTIONS.md. Line numbers go "
            + "stale on the next refactor without anything noticing; identify the exception by "
            + "file and symbol instead.");
    }
}
