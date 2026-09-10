using Aspose.Cells;
using Aspose.Pdf;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using License = Aspose.Slides.License;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     R13-T01: when this suite is asked to run unlicensed, it has to actually be unlicensed.
///     <para>
///         An Aspose licence is process-global. One test that loads one licenses the whole test
///         host, so every test scheduled after it stops exercising evaluation behaviour while
///         still reporting on it. Not hypothetical: <c>LicenseManagerTests</c> called
///         <c>LicenseManager.SetLicense</c> twelve times, and <c>PreflightSizeLimitTests</c> failed
///         two of nine on its own under <c>-SkipLicense</c> but passed all nine when the two
///         classes ran together. Every clean evaluation result reported for this suite was taken
///         from a host in that state.
///     </para>
///     <para>
///         Two guards, because they fail differently. The scan below is deterministic and reads
///         the sources, so which test runs first does not affect it; it is per method rather than
///         per file, because "this file knows about the environment variable" says nothing about
///         whether the test being added today checks it. The probe is a backstop and can only
///         catch contamination that has already happened when it runs — xUnit gives no ordering,
///         so on its own it would miss. Neither replaces the other.
///     </para>
/// </summary>
public class EvaluationHostTests : TestBase
{
    /// <summary>
    ///     The files allowed to load a licence unguarded, and why. <c>TestBase</c> is where the
    ///     suite's own licence handling lives, and it checks the environment variable before it
    ///     reaches any of these calls.
    /// </summary>
    private static readonly string[] MayLoadALicence =
    [
        "Tests/Infrastructure/TestBase.cs"
    ];

    [Fact]
    public void EveryTestThatLoadsALicence_ShouldRefuseToRunOnAnUnlicensedHost()
    {
        var offenders = new List<string>();

        foreach (var (path, text) in TestSources())
        {
            if (MayLoadALicence.Contains(path, StringComparer.Ordinal)) continue;

            var root = CSharpSyntaxTree.ParseText(text).GetRoot();
            foreach (var method in root.DescendantNodes().OfType<MethodDeclarationSyntax>())
            {
                var calls = method.DescendantNodes().OfType<InvocationExpressionSyntax>()
                    .Select(NameOf).ToList();

                // A lambda counts: Record.Exception(() => LicenseManager.SetLicense(config)) loads
                // one just as surely as calling it directly.
                if (!calls.Contains("SetLicense", StringComparer.Ordinal)) continue;
                if (calls.Any(name => name.StartsWith("Skip", StringComparison.Ordinal))) continue;

                offenders.Add($"{path}:{method.Identifier.ValueText}");
            }
        }

        Assert.True(offenders.Count == 0,
            "These tests load an Aspose licence without first checking whether this host was asked "
            + "to stay unlicensed, which licenses the whole process for every test scheduled after "
            + "them: " + string.Join(", ", offenders));
    }

    [SkippableFact]
    public void AnEvaluationRun_ShouldFindEveryComponentUnlicensed()
    {
        Skip.IfNot(TheHostWasAskedToStayUnlicensed(),
            "Only meaningful when the run asked for evaluation mode");

        // Asked of the libraries themselves, not of the environment variable. Reading the variable
        // back would only confirm it is still set, which was never in doubt; the question is
        // whether something loaded a licence in spite of it.
        Assert.False(Document.IsLicensed,
            "Aspose.Pdf is licensed on a host that was asked to stay unlicensed (R13-T01)");
        Assert.False(new Workbook().IsLicensed,
            "Aspose.Cells is licensed on a host that was asked to stay unlicensed (R13-T01)");
        Assert.False(new License().IsLicensed(),
            "Aspose.Slides is licensed on a host that was asked to stay unlicensed (R13-T01)");

        // Words, Email and BarCode expose no such property, so this cannot ask them. The scan
        // above is what covers those: it stops a licence being loaded at all, rather than
        // noticing afterwards that one was.
    }

    /// <summary>The name a call invokes, ignoring whatever it is called on.</summary>
    /// <param name="call">The invocation to name.</param>
    /// <returns>The method name, or an empty string when the expression has no simple name.</returns>
    private static string NameOf(InvocationExpressionSyntax call)
    {
        return call.Expression switch
        {
            MemberAccessExpressionSyntax member => member.Name.Identifier.ValueText,
            IdentifierNameSyntax identifier => identifier.Identifier.ValueText,
            MemberBindingExpressionSyntax binding => binding.Name.Identifier.ValueText,
            _ => string.Empty
        };
    }

    /// <summary>Whether this run asked for evaluation mode.</summary>
    /// <returns><c>true</c> when the environment variable test.ps1 sets is present.</returns>
    private static bool TheHostWasAskedToStayUnlicensed()
    {
        var skip = Environment.GetEnvironmentVariable("SKIP_ASPOSE_LICENSE");
        return string.Equals(skip, "true", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(skip, "1", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Every source file of the test project, with its repository-relative path.</summary>
    /// <returns>Path and content for each file.</returns>
    private static List<(string Path, string Text)> TestSources()
    {
        var root = RepositoryRoot();
        var tests = Path.Combine(root.FullName, "Tests");
        var generated = $"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}";
        var built = $"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}";

        return Directory.EnumerateFiles(tests, "*.cs", SearchOption.AllDirectories)
            .Where(path => !path.Contains(generated, StringComparison.Ordinal)
                           && !path.Contains(built, StringComparison.Ordinal))
            .Select(path => (Path: Path.GetRelativePath(root.FullName, path).Replace('\\', '/'),
                Text: File.ReadAllText(path)))
            .OrderBy(file => file.Path, StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>Walks up from the test binaries to the repository root.</summary>
    /// <returns>The repository root directory.</returns>
    private static DirectoryInfo RepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "AsposeMcpServer.csproj")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        return dir;
    }
}
