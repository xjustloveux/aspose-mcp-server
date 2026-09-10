using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     §22.7 / §23.13.1: the containment rule for a case-sensitive volume, verified without one.
///     <para>
///         The comparison was fixed to <c>OrdinalIgnoreCase</c> on Windows <em>and macOS</em>, on
///         the assumption that macOS volumes fold case. APFS and HFS+ can both be formatted
///         case-sensitive, and there two names differing only in case are two different files — so
///         an ignore-case comparison would place one under a root that does not contain it.
///     </para>
///     <para>
///         The fix stopped inferring volume semantics from the OS name. The evidence was missing
///         because this machine has no case-sensitive volume, but that only made the
///         <em>deployment</em> untestable: the rule is a function of one boolean, so both answers
///         can be exercised anywhere. What a case-sensitive filesystem does with two names is not
///         in question; what this code does with them is.
///     </para>
/// </summary>
public class CleanupPathComparisonTests : TestBase
{
    /// <summary>A rooted path built with this platform's separator.</summary>
    /// <param name="parts">The segments.</param>
    /// <returns>The path.</returns>
    private static string Rooted(params string[] parts)
    {
        return Path.GetFullPath(Path.Combine([Path.GetTempPath(), .. parts]));
    }

    [Fact]
    public void OnACaseSensitivePlatform_APathDifferingOnlyInCase_ShouldNotBeInsideTheRoot()
    {
        // The defect the macOS assumption allowed: `/data/Queue/x` is not under `/data/queue`
        // on such a volume, and treating it as though it were puts a deletion outside the root
        // the queue approved.
        var root = Rooted("cleanup_case_root");
        var elsewhere = Path.Combine(Rooted("cleanup_case_ROOT"), "debt.txt");

        Assert.False(
            CleanupDebtQueue.IsUnderRoot(elsewhere, root, StringComparison.Ordinal),
            "a case-sensitive platform must treat these as two different directories");
    }

    [Fact]
    public void OnACaseFoldingPlatform_APathDifferingOnlyInCase_ShouldBeInsideTheRoot()
    {
        // And the other answer must still hold, or Windows would refuse its own paths.
        var root = Rooted("cleanup_case_root");
        var same = Path.Combine(Rooted("cleanup_case_ROOT"), "debt.txt");

        Assert.True(
            CleanupDebtQueue.IsUnderRoot(same, root, StringComparison.OrdinalIgnoreCase),
            "a case-folding platform must treat these as one directory");
    }

    [Fact]
    public void APathOutsideTheRoot_ShouldBeOutsideUnderEitherRule()
    {
        // The control: the case rule must not be the only thing keeping an unrelated path out.
        var root = Rooted("cleanup_case_root");
        var outside = Path.Combine(Rooted("cleanup_case_other"), "debt.txt");

        foreach (var comparison in new[] { StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase })
            Assert.False(CleanupDebtQueue.IsUnderRoot(outside, root, comparison));
    }

    [Fact]
    public void ARootPrefixThatIsNotADirectoryBoundary_ShouldNotCount()
    {
        // "/data/queue-old/x" starts with "/data/queue" as a string and is not inside it.
        var root = Rooted("cleanup_case_root");
        var sibling = Rooted("cleanup_case_root_old") + Path.DirectorySeparatorChar + "debt.txt";

        foreach (var comparison in new[] { StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase })
            Assert.False(CleanupDebtQueue.IsUnderRoot(sibling, root, comparison),
                "a string prefix is not a directory boundary");
    }

    [Fact]
    public void TheComparisonChosen_ShouldFoldCaseOnlyOnWindows()
    {
        Assert.Equal(StringComparison.OrdinalIgnoreCase, CleanupDebtQueue.ComparisonFor(true));
        Assert.Equal(StringComparison.Ordinal, CleanupDebtQueue.ComparisonFor(false));
    }
}
