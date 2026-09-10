using System.Text.RegularExpressions;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Keeps the number of assertions that cannot fail at zero (NEW-TEST-01).
///     <c>Assert.True(x &gt;= 0)</c> on a count, length or index passes for any result the code
///     could produce, including a wrong one, so a test carrying only that assertion proves the code
///     ran rather than that it worked. All of them were replaced with exact counts, expected
///     identities or expected exceptions, each taken from a measurement rather than a guess, and
///     <see cref="Baseline" /> is now 0: this guard exists to stop new ones appearing.
///     (The summary previously described a non-zero budget that was still being worked down, and
///     referred to a <c>Budget</c> field that no longer exists, R2-C07.)
/// </summary>
public class AssertionQualityGuardTests
{
    /// <summary>
    ///     Exact number of non-falsifiable assertions the suite carries. **Now zero.**
    ///     <para>
    ///         This is a baseline, not a ceiling: the test fails if the count rises <em>or</em>
    ///         falls. A ceiling let the number drift anywhere underneath it, so strengthening an
    ///         assertion left no trace and the figure never had to move.
    ///     </para>
    ///     <para>
    ///         All 158 were replaced on 2026-09-04 with values measured from the code itself:
    ///         each site was briefly instrumented to record what it actually produced, in both
    ///         licensing modes, and then given an exact assertion where the two agreed or a
    ///         measured floor where the evaluation watermark shifts the number. Keeping the
    ///         baseline at zero means any reintroduction fails immediately.
    ///     </para>
    /// </summary>
    private const int Baseline = 0;

    /// <summary>
    ///     Matches an assertion whose condition is satisfied by every value the expression can
    ///     take: a count, length or index compared against zero with <c>&gt;=</c>.
    /// </summary>
    private static readonly Regex NonFalsifiable = new(
        @"Assert\.True\([^;]*?>=\s*0\b[^;]*?\);", RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    /// <summary>
    ///     Walks up from the test binaries to the repository root so the scan works from any
    ///     working directory the runner chooses.
    /// </summary>
    /// <returns>The directory holding the Tests folder.</returns>
    private static DirectoryInfo RepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !Directory.Exists(Path.Combine(dir.FullName, "Tests")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        return dir;
    }

    [Fact]
    public void NonFalsifiableAssertions_ShouldRemainAtZero()
    {
        var testsDirectory = Path.Combine(RepositoryRoot().FullName, "Tests");
        var offenders = new List<string>();
        var scanned = 0;

        foreach (var file in Directory.EnumerateFiles(testsDirectory, "*.cs", SearchOption.AllDirectories))
        {
            if (file.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}") ||
                file.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}") ||
                file.EndsWith(nameof(AssertionQualityGuardTests) + ".cs", StringComparison.Ordinal))
                continue;

            scanned++;
            var lines = File.ReadAllLines(file);
            for (var i = 0; i < lines.Length; i++)
                if (NonFalsifiable.IsMatch(lines[i]))
                    offenders.Add($"{Path.GetFileName(file)}:{i + 1}");
        }

        // A scan that read nothing reports zero offenders, which is the number a clean tree
        // reports too. Without this, a moved root or a changed layout reads as success (R9-T01).
        Assert.True(scanned > 100,
            $"only {scanned} test files were scanned; this guard is reading the wrong tree.");

        Assert.True(offenders.Count == Baseline,
            offenders.Count > Baseline
                ? $"Non-falsifiable assertions rose to {offenders.Count}, above the baseline of "
                  + $"{Baseline}. Assert an exact count, an expected identity or an expected "
                  + $"exception instead. First few: {string.Join(", ", offenders.Take(5))}"
                : $"Non-falsifiable assertions fell to {offenders.Count}, below the baseline of "
                  + $"{Baseline}. That is the intended direction: lower Baseline to "
                  + $"{offenders.Count} so the improvement is locked in and cannot be undone.");
    }
}
