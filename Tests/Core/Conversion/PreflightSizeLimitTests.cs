using System.IO.Compression;
using System.Text;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;
using Document = Aspose.Pdf.Document;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Core.Conversion;

/// <summary>
///     §23.13.1: the size limit applied before the loader, not after it.
///     <para>
///         <c>EnsureModelWithinLimit</c> can only measure a model that exists, so it bounds whether
///         to continue converting — never what the load itself cost. That was recorded for several
///         rounds as needing metadata Aspose does not expose, which was the wrong conclusion:
///         Aspose is not the only thing that can read these formats. A presentation package names
///         one part per slide and a PDF's page tree states its count, both readable without
///         building anything.
///     </para>
///     <para>
///         The preflight is deliberately silent on shapes it cannot read, so those still reach the
///         loaded-model check. Refusing on a number it could not establish would turn a cheap
///         optimisation into false refusals — which is why the "cannot read it" cases below matter
///         as much as the refusals.
///     </para>
/// </summary>
[Collection("SerialSlides")]
public class PreflightSizeLimitTests : TestBase
{
    /// <summary>Writes a presentation with the given number of slides.</summary>
    /// <param name="name">Fixture file name.</param>
    /// <param name="slides">How many slides to add beyond the default one.</param>
    /// <returns>The file path.</returns>
    private string APresentation(string name, int slides)
    {
        var path = CreateTestFilePath(name);

        using var slidesGate = SlidesGate.Enter();
        using var presentation = new Presentation();
        for (var i = 1; i < slides; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);

        presentation.Save(path, SaveFormat.Pptx);
        return path;
    }

    [Fact]
    public void ThePreflight_ShouldCountSlidesWithoutOpeningThePresentation()
    {
        var path = APresentation("preflight_count.pptx", 5);

        Assert.Equal(5, DocumentSizePreflight.SlideCount(path));
    }

    [Fact]
    public void AFileOverTheSlideLimit_ShouldBeRefusedBeforeTheLoader()
    {
        var path = APresentation("preflight_over.pptx", 6);
        var limits = new InMemoryModelLimits(PowerPointSlides: 3);

        var refusal = Assert.Throws<ArgumentException>(() =>
            DocumentConverter.EnsureFileWithinLimit(path, DocumentType.PowerPoint, limits));

        Assert.Contains("6", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("slides", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AFileWithinTheLimit_ShouldNotBeRefused()
    {
        var path = APresentation("preflight_under.pptx", 2);
        var limits = new InMemoryModelLimits(PowerPointSlides: 10);

        var exception = Record.Exception(() =>
            DocumentConverter.EnsureFileWithinLimit(path, DocumentType.PowerPoint, limits));

        Assert.Null(exception);
    }

    [Fact]
    public void ADisabledLimit_ShouldNotBeApplied()
    {
        var path = APresentation("preflight_disabled.pptx", 6);
        var limits = new InMemoryModelLimits(PowerPointSlides: null);

        var exception = Record.Exception(() =>
            DocumentConverter.EnsureFileWithinLimit(path, DocumentType.PowerPoint, limits));

        Assert.Null(exception);
    }

    [Fact]
    public void AFileThePreflightCannotRead_ShouldNotBeRefused()
    {
        // The property that keeps this from becoming a source of false refusals: a shape it does
        // not understand answers "unknown", and the loaded-model check remains the only bound.
        var notAnArchive = CreateTestFilePath("preflight_unreadable.pptx");
        File.WriteAllText(notAnArchive, "this is not a package");

        Assert.Null(DocumentSizePreflight.SlideCount(notAnArchive));

        var exception = Record.Exception(() =>
            DocumentConverter.EnsureFileWithinLimit(notAnArchive, DocumentType.PowerPoint,
                new InMemoryModelLimits(PowerPointSlides: 1)));

        Assert.Null(exception);
    }

    [Fact]
    public void AZipThatIsNotAPresentation_ShouldAnswerUnknownRatherThanZero()
    {
        // Reporting zero would read as "well within the limit" for any archive at all.
        var path = CreateTestFilePath("preflight_other.zip");
        using (var archive = ZipFile.Open(
                   path, ZipArchiveMode.Create))
        {
            archive.CreateEntry("word/document.xml");
        }

        Assert.Null(DocumentSizePreflight.SlideCount(path));
    }

    /// <summary>Writes a valid PDF with the given number of pages, built without any loader.</summary>
    /// <param name="name">Fixture file name.</param>
    /// <param name="pages">How many page objects the tree holds.</param>
    /// <param name="before">Bytes of body content written ahead of the page tree.</param>
    /// <param name="after">Bytes of body content written behind it.</param>
    /// <returns>The file path.</returns>
    /// <remarks>
    ///     Hand-written on purpose. The preflight's whole claim is that it reads a PDF without
    ///     Aspose, so building the fixture with Aspose measured the two against each other and
    ///     imported the evaluation-mode page cap along the way — which is why these cases could
    ///     only ever run licensed (R13-T01). Filler on either side is what lets a fixture place the
    ///     page tree where it wants it in the file.
    /// </remarks>
    private string APdf(string name, int pages, int before = 0, int after = 0)
    {
        var path = CreateTestFilePath(name);
        var objects = new List<string> { "<< /Type /Catalog /Pages 2 0 R >>" };

        if (before > 0) objects.Add(Filler(before));

        var tree = objects.Count + 1;
        var kids = string.Join(" ", Enumerable.Range(0, pages).Select(i => $"{tree + 1 + i} 0 R"));
        objects.Add($"<< /Type /Pages /Kids [{kids}] /Count {pages} >>");
        for (var i = 0; i < pages; i++)
            objects.Add($"<< /Type /Page /Parent {tree} 0 R /MediaBox [0 0 612 792] >>");

        if (after > 0) objects.Add(Filler(after));

        // The catalogue names the tree, which is no longer always object 2.
        objects[0] = $"<< /Type /Catalog /Pages {tree} 0 R >>";

        var pdf = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        for (var i = 0; i < objects.Count; i++)
        {
            offsets.Add(pdf.Length);
            pdf.Append($"{i + 1} 0 obj\n{objects[i]}\nendobj\n");
        }

        var startxref = pdf.Length;
        pdf.Append($"xref\n0 {objects.Count + 1}\n0000000000 65535 f \n");
        foreach (var offset in offsets) pdf.Append($"{offset:D10} 00000 n \n");
        pdf.Append($"trailer\n<< /Size {objects.Count + 1} /Root 1 0 R >>\n");
        pdf.Append($"startxref\n{startxref}\n%%EOF\n");

        File.WriteAllBytes(path, Encoding.Latin1.GetBytes(pdf.ToString()));
        return path;
    }

    /// <summary>A body object of the given size, holding nothing the preflight looks for.</summary>
    /// <param name="bytes">How large the stream is.</param>
    /// <returns>The object's text.</returns>
    private static string Filler(int bytes)
    {
        return $"<< /Length {bytes} >>\nstream\n{new string('A', bytes)}\nendstream";
    }

    [Fact]
    public void ThePreflight_ShouldCountPdfPagesFromThePageTree()
    {
        Assert.Equal(4, DocumentSizePreflight.PageCount(APdf("preflight_pages.pdf", 4)));
    }

    [Fact]
    public void APdfWhosePageTreeIsFarFromTheEnd_ShouldStillBeCounted()
    {
        // The case the old fixture could not reach. The preflight read only the last 64 KB, on the
        // assumption that the catalogue is written last; the producer this server itself uses puts
        // it at byte 157, so as soon as a document grew past that window its page tree was outside
        // the only part being read and every such file answered "unknown" (R13-C01). The fixture
        // meant to cover this was four pages and a few hundred bytes, so it never left the window.
        var path = APdf("preflight_pages_padded.pdf", 4, after: 80 * 1024);

        Assert.True(new FileInfo(path).Length > 64 * 1024, "the fixture must exceed the old window");
        Assert.Equal(4, DocumentSizePreflight.PageCount(path));
    }

    [Fact]
    public void APdfWhosePageTreeIsBeyondBothWindows_ShouldAnswerUnknown()
    {
        // Still a scan, not a parser. A tree that falls in neither window is not found, and saying
        // so is what keeps a count this could not establish from becoming a refusal.
        var path = APdf("preflight_pages_middle.pdf", 4, 300 * 1024, 300 * 1024);

        Assert.Null(DocumentSizePreflight.PageCount(path));

        DocumentConverter.EnsureFileWithinLimit(path, DocumentType.Pdf,
            new InMemoryModelLimits(PdfPages: 1));
    }

    [Fact]
    public void AnInteriorNodeOfThePageTree_ShouldNotBeReadAsTheDocumentsCount()
    {
        // A large document's tree has more than one level, and every node states how many pages
        // hang below it. Reading whichever is found first reports a fraction of the document and
        // the limit quietly stops applying; only the root, the node with no /Parent, is the total.
        var path = CreateTestFilePath("preflight_pages_interior.pdf");
        File.WriteAllText(path,
            "%PDF-1.7\n"
            + "1 0 obj\n<< /Type /Pages /Parent 9 0 R /Kids [4 0 R] /Count 2 >>\nendobj\n"
            + "2 0 obj\n<< /Type /Pages /Kids [1 0 R 3 0 R] /Count 40 >>\nendobj\n"
            + "%%EOF\n");

        Assert.Equal(40, DocumentSizePreflight.PageCount(path));
    }

    [Fact]
    public void TwoPageTreesDisagreeing_ShouldAnswerUnknownRatherThanPickOne()
    {
        // An incrementally updated file carries the superseded tree as well, and nothing here can
        // tell which is current. Answering the larger would refuse a document that is within its
        // limit, which is the one outcome this preflight must never produce.
        var path = CreateTestFilePath("preflight_pages_revised.pdf");
        File.WriteAllText(path,
            "%PDF-1.7\n"
            + "1 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 2 >>\nendobj\n"
            + "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 900 >>\nendobj\n"
            + "%%EOF\n");

        Assert.Null(DocumentSizePreflight.PageCount(path));
    }

    [SkippableFact]
    public void APdfSavedByTheProducerThisServerUses_ShouldBeCounted()
    {
        // The hand-written fixtures assume a shape; this one is the shape Aspose actually writes.
        // It cannot be built unlicensed — evaluation mode refuses past the fourth page — so it says
        // so rather than failing, and the deterministic cases above carry the coverage there.
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Evaluation mode caps a document at four pages");

        var path = CreateTestFilePath("preflight_pages_aspose.pdf");
        var pdf = new Document();
        for (var i = 0; i < 5; i++) pdf.Pages.Add();
        pdf.Save(path);

        Assert.Equal(5, DocumentSizePreflight.PageCount(path));
    }

    [Fact]
    public void APdfOverThePageLimit_ShouldBeRefusedBeforeTheLoader()
    {
        var path = APdf("preflight_pages_over.pdf", 6);

        var refusal = Assert.Throws<ArgumentException>(() =>
            DocumentConverter.EnsureFileWithinLimit(path, DocumentType.Pdf,
                new InMemoryModelLimits(PdfPages: 2)));

        Assert.Contains("pages", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AFormatWithNoPreflight_ShouldPassThroughSilently()
    {
        // Word and Excel have no cheap count here; they must not be refused for lack of one.
        var path = CreateTestFilePath("preflight_word.docx");
        var document = new Aspose.Words.Document();
        new DocumentBuilder(document).Writeln("body");
        document.Save(path);

        var exception = Record.Exception(() =>
            DocumentConverter.EnsureFileWithinLimit(path, DocumentType.Word,
                new InMemoryModelLimits(1)));

        Assert.Null(exception);
    }
}
