using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

/// <summary>
///     Header and footer text must not be able to carry an active field code (R2-C08, from the
///     §13.7 deferred list).
///     <para>
///         Header text is parsed for <c>{...}</c> tokens, and an unrecognised token was passed
///         straight to <c>DocumentBuilder.InsertField</c>. Measured against the pinned
///         Aspose.Words: an <c>INCLUDETEXT</c> field does not resolve on save or on PDF render, but
///         it does resolve on <c>Document.UpdateFields()</c> — which this server calls in five
///         places, one of them <c>GetHeadersFootersHandler</c>, the read side of the very path that
///         writes the header. That turns "set a header" followed by "read the headers" into an
///         arbitrary local file read that never passes the path allowlist, because the path lives
///         inside a field code rather than in a path parameter.
///     </para>
/// </summary>
public class HeaderFieldCodeInjectionTests : WordTestBase
{
    private readonly WordHeaderFooterTool _tool;

    /// <summary>Builds the tool under test.</summary>
    public HeaderFieldCodeInjectionTests()
    {
        _tool = new WordHeaderFooterTool();
    }

    [Fact]
    public void SetHeaderText_WithAnIncludeTextField_ShouldNotReadTheFile()
    {
        var secretPath = Path.Combine(TestDir, "not-for-the-caller.txt");
        System.IO.File.WriteAllText(secretPath, "CANARY_STRING_12345");

        var docPath = CreateTestFilePath("header_field_injection.docx");
        var doc = new Document();
        new DocumentBuilder(doc).Writeln("body");
        doc.Save(docPath);

        var fieldCode = "{INCLUDETEXT \"" + secretPath.Replace("\\", "\\\\") + "\"}";

        // Either the write is refused, or it is inert. Both are acceptable; silently creating a
        // live field is not.
        try
        {
            _tool.Execute("set_header", docPath, outputPath: docPath, headerCenter: fieldCode);
        }
        catch (ArgumentException ex)
        {
            // A refusal only counts when it is this refusal. Accepting any ArgumentException let
            // an unrelated failure earlier in the call - a bad path, a missing file - end the test
            // green while the field policy itself was never exercised (R7-T01).
            Assert.Contains("Unsupported field code", ex.Message, StringComparison.Ordinal);
            Assert.Contains("INCLUDETEXT", ex.Message, StringComparison.OrdinalIgnoreCase);
            return;
        }

        var written = new Document(docPath);
        written.UpdateFields();

        Assert.DoesNotContain("CANARY_STRING_12345", written.GetText());
    }

    [Fact]
    public void SetHeaderText_OnAMissingFile_ShouldFailForADifferentReason()
    {
        // The negative half of the fixture above: the two failures must not be confusable, or the
        // injection test could pass on the strength of an error that has nothing to do with fields.
        var missing = Path.Combine(TestDir, "no-such-document.docx");

        var ex = Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("set_header", missing, outputPath: missing, headerCenter: "plain text"));

        Assert.DoesNotContain("Unsupported field code", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void SetHeaderText_WithADocumentedField_ShouldStillWork()
    {
        var docPath = CreateTestFilePath("header_field_supported.docx");
        var doc = new Document();
        new DocumentBuilder(doc).Writeln("body");
        doc.Save(docPath);

        _tool.Execute("set_header", docPath, outputPath: docPath,
            headerCenter: "Page {PAGE} of {NUMPAGES}");

        var written = new Document(docPath);
        var header = written.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        Assert.Contains(header.Range.Fields,
            f => f.Type == FieldType.FieldPage);
    }
}
