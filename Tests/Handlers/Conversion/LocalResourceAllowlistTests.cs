using System.Text;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R17-M01: whether the pinned renderer reads a local file the allowlist does not cover.
///     <para>
///         The external-resource scan is about <em>remote</em> references — it exists so converting
///         a document cannot make the server fetch from hosts the caller could not otherwise reach.
///         A relative reference that resolves to a file on this disk is a different question, and
///         §28.13 recorded it as unanswered: the scanner skips those, and nothing had established
///         whether Aspose then reads them. An unanswered question about a renderer is not evidence
///         either way, so this asks it.
///     </para>
///     <para>
///         Answered by conversion rather than by reading the scanner, because what the scanner does
///         is not what decides this. If the renderer reads the file, the same allowlist that bounds
///         every other read has to be applied before it is handed the document; if it does not, this
///         stays as the regression fixture that would notice a version that starts to.
///     </para>
///     <para>
///         Each case pins the exception it expects. The first version accepted any failure and
///         returned, so a conversion that broke for an unrelated reason produced the same green as
///         a guard doing its job (R18-TEST01) — the shape this suite has been correcting all round.
///     </para>
/// </summary>
public class LocalResourceAllowlistTests : TestBase
{
    /// <summary>Text that could only appear in the output by being read off disk.</summary>
    private const string Sentinel = "SENTINEL7f3a2b9c0d4e";

    /// <summary>Converts one document and reports what stopped it, if anything.</summary>
    /// <param name="input">The document to convert.</param>
    /// <param name="output">Where the PDF goes.</param>
    /// <param name="allowed">The allowlist to convert under.</param>
    /// <param name="allowExternalResources">Whether referenced resources may be loaded.</param>
    /// <returns>The exception's type name, or null when the conversion completed.</returns>
    private static string? Convert(string input, string output, string[] allowed,
        bool allowExternalResources)
    {
        try
        {
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, allowed,
                allowExternalResources);
            return null;
        }
        catch (Exception exception)
        {
            return exception.GetType().Name;
        }
    }

    /// <summary>Writes a document referencing another file, and the file it references.</summary>
    /// <param name="directory">Where the document goes.</param>
    /// <param name="reference">The reference as it appears in the markup.</param>
    /// <param name="target">Where the referenced file goes.</param>
    /// <returns>The document's path.</returns>
    private static string ADocumentReferencing(string directory, string reference, string target)
    {
        Directory.CreateDirectory(directory);
        Directory.CreateDirectory(Path.GetDirectoryName(target)!);
        File.WriteAllText(target, $"<html><body><p>{Sentinel}</p></body></html>", Encoding.UTF8);

        var input = Path.Combine(directory, "document.html");
        File.WriteAllText(input,
            "<html><body><p>the document itself</p>"
            + $"<iframe src=\"{reference}\"></iframe></body></html>",
            Encoding.UTF8);

        return input;
    }

    [SkippableFact]
    public void UnderTheDefaultPolicy_TheRendererReadsNoReferencedFileAtAll()
    {
        // The measurement §28.13 asked for. `allowExternalResources: false` is the default, and
        // Aspose.Pdf 23.10 refuses every referenced resource under it — it does not distinguish a
        // relative local path from a remote URL. So the shape that was recorded as "the scanner
        // skips these and nobody has checked whether the renderer reads them" does not arise:
        // nothing is read.
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "The output is rewritten unlicensed");

        var allowed = Path.Combine(TestDir, "default_policy");
        var input = ADocumentReferencing(allowed, "../outside/secret.html",
            Path.Combine(TestDir, "outside", "secret.html"));
        var output = Path.Combine(allowed, "document.pdf");

        Assert.Equal("UnauthorizedAccessException", Convert(input, output, [allowed], false));
        Assert.False(File.Exists(output), "a refused conversion still produced a document");
    }

    [SkippableFact]
    public void AReferenceInsideTheAllowlist_IsRefusedByTheSamePolicy()
    {
        // Recorded because it is the reason the case above is not a gap being papered over: the
        // renderer's resource loading is off wholesale, so a caller who wants a local include has
        // to opt in, and opting in is what the next case is about.
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "The output is rewritten unlicensed");

        var allowed = Path.Combine(TestDir, "inside_default");
        var input = ADocumentReferencing(allowed, "part.html", Path.Combine(allowed, "part.html"));

        Assert.Equal("UnauthorizedAccessException",
            Convert(input, Path.Combine(allowed, "document.pdf"), [allowed], false));
    }

    [SkippableFact]
    public void WithExternalResourcesAllowed_AReferenceOutsideTheAllowlistShouldNotReachTheOutput()
    {
        // The case with teeth. `--allow-external-resources` is a deployment decision about reaching
        // the network; it is not a decision to let a converted document read any file on the disk.
        // The allowlist bounds every other read this server performs and has to bound this one.
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "The output is rewritten unlicensed");

        var allowed = Path.Combine(TestDir, "opted_in");
        var input = ADocumentReferencing(allowed, "../beyond/secret.html",
            Path.Combine(TestDir, "beyond", "secret.html"));
        var output = Path.Combine(allowed, "document.pdf");

        var failure = Convert(input, output, [allowed], true);

        // Pinned rather than "any failure will do". Accepting every exception meant a conversion
        // that broke for an unrelated reason looked exactly like a guard doing its job
        // (R18-TEST01). Measured against the pinned Aspose.PDF.Drawing 23.10.0: it refuses.
        Assert.Equal("UnauthorizedAccessException", failure);
        Assert.False(File.Exists(output), "a refused conversion still produced a document");
    }
}
