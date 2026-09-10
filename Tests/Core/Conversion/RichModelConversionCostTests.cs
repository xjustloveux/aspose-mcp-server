using Aspose.Cells;
using Aspose.Pdf.Text;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.Conversion;

/// <summary>
///     R9-C02: what a conversion costs, measured on content that carries what documents carry.
///     <para>
///         The first round of measurements priced the cheapest shape of each format — plain
///         paragraphs, short strings, empty slides, text-only pages — and set the limits below what
///         those numbers allowed, on the reasoning that richer content would cost more. It does,
///         but by factors that differ by more than an order of magnitude between formats: a node
///         carrying a table costs almost the same as a plain one, while a slide carrying a dozen
///         text shapes costs forty times an empty one. A single safety margin could not stand in
///         for that, so each limit is now derived from its own expensive measurement.
///     </para>
///     <para>
///         The second thing measured here is the file path, which the refusal points at. It is
///         cheaper, but it is not free: the source model is held either way, and only the second
///         copy of the result is avoided. The refusal used to say "convert to a file instead",
///         which reads as though the cost went away.
///     </para>
///     <para>
///         <b>What these numbers are, and are not.</b> They are the increase in <em>managed heap</em>
///         during a conversion, sampled after the source model has already been built. They do not
///         include the model itself, Aspose native allocations, or the process working set, and
///         each format is measured in-memory first and to a file second, so an ordering or cache
///         effect is not excluded. They are therefore sound for comparing the two conversion paths
///         against each other — which is what they are used for — and are not a statement of the
///         full per-unit cost of a model. Establishing that needs an isolated process observing
///         private bytes from before the model is built, which this suite does not do (§21.7).
///     </para>
///     <para>
///         Each measurement builds a model larger than evaluation mode will construct, so each is
///         skipped without the matching licence rather than failing there. Measured in a full run
///         and failing in isolation is not a passing test — it is an order-dependent one (§21.8).
///     </para>
/// </summary>
[Collection("SerialSlides")]
public class RichModelConversionCostTests : TestBase
{
    /// <summary>
    ///     The heap-delta figure the defaults were chosen against, per format.
    ///     <para>
    ///         Not a memory budget and not a process bound: it is the scale the extrapolation
    ///         below is compared with, so a default raised far past the fixture it came from shows
    ///         up as a red test. Calling it a target invited the reading that a conversion is held
    ///         under it, which nothing here measures (§23.7).
    ///     </para>
    /// </summary>
    private const long ReferenceHeapDeltaBytes = 256L * 1024 * 1024;

    /// <summary>
    ///     Headroom on the target. Heap sampling is not exact and a limit derived from one
    ///     measurement should not fail on the noise of the next.
    /// </summary>
    private const double AllowedOvershoot = 1.35;

    /// <summary>
    ///     Runs an action while sampling the managed heap, and reports the highest reading.
    /// </summary>
    /// <param name="action">The work to measure.</param>
    /// <returns>The peak managed heap above the starting point, in bytes.</returns>
    private static long MeasurePeakHeap(Action action)
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var baseline = GC.GetTotalMemory(true);
        var peak = 0L;
        using var sampling = new CancellationTokenSource();

        var sampler = Task.Run(() =>
        {
            while (!sampling.IsCancellationRequested)
            {
                var current = GC.GetTotalMemory(false) - baseline;
                if (current > peak) Interlocked.Exchange(ref peak, current);
                Thread.Sleep(5);
            }
        });

        try
        {
            action();
        }
        finally
        {
            sampling.Cancel();
            sampler.Wait(TimeSpan.FromSeconds(5));
        }

        return Math.Max(peak, GC.GetTotalMemory(false) - baseline);
    }

    /// <summary>The conversion options a measurement uses.</summary>
    /// <returns>Options allowing writes into the test directory, with no model limit.</returns>
    private ConversionOptions Unlimited()
    {
        return new ConversionOptions
        {
            RecoveryDirectory = Path.GetTempPath(),
            AllowedBasePaths = [TestDir],
            ModelLimits = new InMemoryModelLimits(null, null, null, null)
        };
    }

    /// <summary>
    ///     Asserts that the format's default limit, priced at the cost just measured, stays within
    ///     the peak this project targets.
    /// </summary>
    /// <param name="format">The format, for the failure message.</param>
    /// <param name="peakBytes">The measured peak.</param>
    /// <param name="units">How many units of the model produced it.</param>
    /// <param name="limit">The default limit for that format.</param>
    private static void TheDefaultLimitShouldStayNearItsFixture(
        string format, long peakBytes, int units, int limit)
    {
        var perUnit = (double)peakBytes / units;
        var atTheLimit = perUnit * limit;

        // A heuristic, and named as one. This multiplies a *managed-heap increment measured on
        // one fixture* out to the default limit; it does not include the model itself, native
        // allocations or the process working set, so it cannot say what a conversion at the limit
        // really costs. What it can do is notice when a default is raised far past the shape it
        // was derived from — which is the regression worth catching here (§22.8).
        Assert.True(atTheLimit <= ReferenceHeapDeltaBytes * AllowedOvershoot,
            $"{format}: {perUnit / 1024:F1} KB per unit on this fixture x the default limit of "
            + $"{limit:N0} extrapolates to {atTheLimit / 1048576:F0} MB, past the "
            + $"{ReferenceHeapDeltaBytes / 1048576} MB these defaults were derived against. This is a "
            + "heuristic on one measured shape, not a bound on real memory: lower the default, "
            + "or re-derive it from a fixture that justifies the new figure.");
    }

    /// <summary>
    ///     Asserts the file path carries a real share of the same cost, which is why the refusal
    ///     may not claim it avoids the problem.
    /// </summary>
    /// <param name="format">The format, for the failure message.</param>
    /// <param name="memoryPeak">Peak converting in memory.</param>
    /// <param name="filePeak">Peak converting to a file.</param>
    private static void TheFilePathShouldStillHoldTheModel(
        string format, long memoryPeak, long filePeak)
    {
        Assert.True(filePeak > memoryPeak / 5,
            $"{format}: converting to a file peaked at {filePeak / 1048576} MB against "
            + $"{memoryPeak / 1048576} MB in memory. If the file path really cost almost nothing, "
            + "the refusal could point at it as an escape — this assertion exists to notice if "
            + "that ever becomes true.");
    }

    [SkippableFact]
    public void AWordDocumentOfTables_ShouldCostWhatTheWordLimitAssumes()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words,
            "A rich model is larger than evaluation mode will build, so this measures nothing there.");

        var document = new Document();
        var builder = new DocumentBuilder(document);
        for (var row = 0; row < 1500; row++)
        {
            builder.StartTable();
            for (var cell = 0; cell < 4; cell++)
            {
                builder.InsertCell();
                builder.Write(new string('x', 64));
            }

            builder.EndRow();
            builder.EndTable();
        }

        var nodes = document.GetChildNodes(NodeType.Any, true).Count;

        var memory = MeasurePeakHeap(() =>
        {
            using var stream = DocumentConverter.ConvertToStream(
                document, DocumentType.Word, "pdf", Unlimited());
        });
        var file = MeasurePeakHeap(() => DocumentConverter.ConvertWordDocument(
            document, CreateTestFilePath("rich_word.pdf"), "pdf", null, Unlimited()));

        TheDefaultLimitShouldStayNearItsFixture("Word", memory, nodes,
            new InMemoryModelLimits().WordNodes!.Value);
        TheFilePathShouldStillHoldTheModel("Word", memory, file);
    }

    [SkippableFact]
    public void AWorkbookOfFormulas_ShouldCostWhatTheExcelLimitAssumes()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells,
            "A rich model is larger than evaluation mode will build, so this measures nothing there.");

        var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;
        const int rows = 4000;
        for (var r = 0; r < rows; r++)
        for (var c = 0; c < 8; c++)
            if (c == 7) cells[r, c].Formula = $"=SUM(A{r + 1}:G{r + 1})";
            else cells[r, c].PutValue(r * 8 + c);

        var memory = MeasurePeakHeap(() =>
        {
            using var stream = DocumentConverter.ConvertToStream(
                workbook, DocumentType.Excel, "pdf", Unlimited());
        });
        var file = MeasurePeakHeap(() => DocumentConverter.ConvertExcelDocument(
            workbook, CreateTestFilePath("rich_excel.pdf"), "pdf", null, Unlimited()));

        TheDefaultLimitShouldStayNearItsFixture("Excel", memory, rows * 8,
            new InMemoryModelLimits().ExcelCells!.Value);
        TheFilePathShouldStillHoldTheModel("Excel", memory, file);
    }

    [SkippableFact]
    public void APresentationOfShapes_ShouldCostWhatThePowerPointLimitAssumes()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides,
            "A rich model is larger than evaluation mode will build, so this measures nothing there.");

        using var gate = SlidesGate.Enter();
        using var presentation = new Presentation();

        for (var slide = 0; slide < 120; slide++)
        {
            var added = presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            for (var shape = 0; shape < 12; shape++)
                added.Shapes.AddAutoShape(ShapeType.Rectangle,
                        10 + shape * 5, 10 + shape * 5, 80, 40)
                    .TextFrame.Text = new string('x', 200);
        }

        var slides = presentation.Slides.Count;

        var memory = MeasurePeakHeap(() =>
        {
            using var stream = DocumentConverter.ConvertToStream(
                presentation, DocumentType.PowerPoint, "pdf", Unlimited());
        });
        var file = MeasurePeakHeap(() => DocumentConverter.ConvertPowerPointDocument(
            presentation, CreateTestFilePath("rich_ppt.pdf"), "pdf", null, Unlimited()));

        TheDefaultLimitShouldStayNearItsFixture("PowerPoint", memory, slides,
            new InMemoryModelLimits().PowerPointSlides!.Value);
        TheFilePathShouldStillHoldTheModel("PowerPoint", memory, file);
    }

    [SkippableFact]
    public void APdfOfTextFragments_ShouldCostWhatThePdfLimitAssumes()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf,
            "A rich model is larger than evaluation mode will build, so this measures nothing there.");

        var pdf = new Aspose.Pdf.Document();
        const int pages = 300;
        for (var i = 0; i < pages; i++)
        {
            var page = pdf.Pages.Add();
            page.Paragraphs.Add(new TextFragment(new string('x', 2000)));
            page.Paragraphs.Add(new TextFragment(new string('y', 2000)));
        }

        var memory = MeasurePeakHeap(() =>
        {
            using var stream = DocumentConverter.ConvertToStream(
                pdf, DocumentType.Pdf, "docx", Unlimited());
        });
        var file = MeasurePeakHeap(() => DocumentConverter.ConvertPdfDocument(
            pdf, CreateTestFilePath("rich_pdf.docx"), "docx", Unlimited()));

        TheDefaultLimitShouldStayNearItsFixture("PDF", memory, pdf.Pages.Count,
            new InMemoryModelLimits().PdfPages!.Value);
        TheFilePathShouldStillHoldTheModel("PDF", memory, file);
    }

    [Fact]
    public void TheRefusalMessage_ShouldNotClaimTheFilePathAvoidsTheCost()
    {
        // The wording is load-bearing: it is the only thing that tells a refused caller what to do,
        // and "instead" read as though the memory problem went away on that path.
        var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;
        for (var r = 0; r < 40; r++)
        for (var c = 0; c < 8; c++)
            cells[r, c].PutValue(r);

        var refusal = Assert.Throws<ArgumentException>(() =>
        {
            using var stream = DocumentConverter.ConvertToStream(
                workbook, DocumentType.Excel, "pdf",
                new ConversionOptions
                    { RecoveryDirectory = Path.GetTempPath(), ModelLimits = new InMemoryModelLimits(ExcelCells: 10) });
        });

        Assert.Contains("costs less memory", refusal.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("instead", refusal.Message, StringComparison.Ordinal);
    }
}
