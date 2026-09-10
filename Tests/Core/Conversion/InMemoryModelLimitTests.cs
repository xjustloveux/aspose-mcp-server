using Aspose.Words;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.Conversion;

/// <summary>
///     §19.10.3: the one thing that decides peak memory during an in-memory conversion is the size
///     of the loaded document, and neither byte cap could see it.
///     <para>
///         <see cref="RenderBudget.MaxInMemoryOutputBytes" /> bounds what is handed back and
///         <c>SessionConfig.MaxFileSizeMb</c> bounds what is read off disk; measured, peak heap
///         tracked the decompressed model at roughly 2,400x the compressed input. So the mechanism
///         to refuse an over-large document exists now, with defaults taken from peak-heap
///         measurements at three sizes per format and set below what those measurements would
///         allow — the shapes priced are the cheapest of their kind. Any limit can be turned off,
///         and the refusal points at the file path, which streams and is exempt.
///     </para>
/// </summary>
public class InMemoryModelLimitTests : TestBase
{
    /// <summary>Builds a Word document with a known number of paragraphs.</summary>
    /// <param name="paragraphs">How many paragraphs to write.</param>
    /// <returns>The document.</returns>
    private static Document ADocument(int paragraphs)
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);
        for (var i = 0; i < paragraphs; i++) builder.Writeln("line");

        return document;
    }

    [Fact]
    public void AnOrdinaryDocument_ShouldConvertUnderTheDefaults()
    {
        // The defaults have to leave ordinary work alone, or the limit is just a broken feature.
        var document = ADocument(200);

        var bytes = DocumentConverter.ConvertToBytes(document, DocumentType.Word, "pdf");

        Assert.NotEmpty(bytes);
    }

    /// <param name="unit">What the limit counts.</param>
    /// <param name="value">The default for that format.</param>
    [Theory]
    [InlineData("Word nodes", 60_000)]
    [InlineData("Excel cells", 120_000)]
    [InlineData("PowerPoint slides", 600)]
    [InlineData("PDF pages", 2_000)]
    public void EveryFormat_ShouldCarryTheMeasuredDefault(string unit, int value)
    {
        // Recorded so a change to any of them has to be deliberate. The numbers come from peak-heap
        // measurements at three sizes per format (§19.10.3), set below what the measurement would
        // allow because the shapes measured — plain text, short strings, empty slides, text-only
        // pages — are the cheapest of their kind.
        var limits = ConversionOptions.WithoutAHost().ModelLimits;
        var actual = unit switch
        {
            "Word nodes" => limits.WordNodes,
            "Excel cells" => limits.ExcelCells,
            "PowerPoint slides" => limits.PowerPointSlides,
            _ => limits.PdfPages
        };

        Assert.Equal(value, actual);
    }

    [Fact]
    public void TurningALimitOff_ShouldBePossible()
    {
        // An operator whose documents are larger than the defaults must be able to say so rather
        // than be told their work is impossible.
        var document = ADocument(200);
        var options = new ConversionOptions
            { RecoveryDirectory = Path.GetTempPath(), ModelLimits = new InMemoryModelLimits(null) };

        using var stream = DocumentConverter.ConvertToStream(document, DocumentType.Word, "pdf", options);

        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void WithALimitConfigured_ALargerDocument_ShouldBeRefusedBeforeAnythingIsRendered()
    {
        var document = ADocument(200);
        var nodes = document.GetChildNodes(NodeType.Any, true).Count;

        var options = new ConversionOptions
        {
            RecoveryDirectory = Path.GetTempPath(),
            ModelLimits = new InMemoryModelLimits(nodes - 1)
        };

        var refusal = Assert.Throws<ArgumentException>(() =>
            DocumentConverter.ConvertToStream(document, DocumentType.Word, "pdf", options));

        Assert.Contains("converts in memory", refusal.Message, StringComparison.Ordinal);
        Assert.Contains(nodes.ToString("N0"), refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void WithALimitConfigured_ADocumentInsideIt_ShouldStillConvert()
    {
        // A guard that refuses everything is no better than one that refuses nothing.
        var document = ADocument(200);
        var nodes = document.GetChildNodes(NodeType.Any, true).Count;

        var options = new ConversionOptions
        {
            RecoveryDirectory = Path.GetTempPath(),
            ModelLimits = new InMemoryModelLimits(nodes + 1)
        };

        using var stream = DocumentConverter.ConvertToStream(document, DocumentType.Word, "pdf", options);

        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ALimitForAnotherFormat_ShouldNotApplyToThisOne()
    {
        // The units differ per model, so the limits do too; a slide count cannot speak for a Word
        // document.
        var document = ADocument(200);
        var options = new ConversionOptions
        {
            RecoveryDirectory = Path.GetTempPath(),
            ModelLimits = new InMemoryModelLimits(PowerPointSlides: 1)
        };

        using var stream = DocumentConverter.ConvertToStream(document, DocumentType.Word, "pdf", options);

        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void TheDiskPath_ShouldNotBeBoundedByAnInMemoryLimit()
    {
        // Writing a file streams the result, so the reason for the limit does not apply there.
        // Binding the two together would refuse conversions that never held the document twice.
        var document = ADocument(200);
        var nodes = document.GetChildNodes(NodeType.Any, true).Count;
        var output = CreateTestFilePath("model_limit_disk.pdf");

        DocumentConverter.ConvertWordDocument(document, output, "pdf", null, new ConversionOptions
        {
            RecoveryDirectory = Path.GetTempPath(),
            ModelLimits = new InMemoryModelLimits(nodes - 1)
        });

        Assert.True(File.Exists(output));
    }
}
