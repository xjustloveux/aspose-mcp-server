using Aspose.Words;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.Conversion;

/// <summary>
///     §17.4.2 / R8-C03: what an in-memory conversion at the budget actually costs, measured rather
///     than assumed.
///     <para>
///         <see cref="RenderBudget.MaxInMemoryOutputBytes" /> was set to 64 MiB as a reasoned
///         reduction from the 2 GiB on-disk limit, with the reasoning that a response built in
///         memory and handed back as an array is held more than once at the moment of the copy. The
///         number had never been checked against a real conversion, which is what these do: they
///         record the peak managed allocation for a conversion at the limit, and assert the
///         relationship the limit exists to bound rather than a hard-coded megabyte count that
///         would break on any allocator change.
///     </para>
/// </summary>
[Collection("SerialSlides")]
public class InMemoryBudgetMeasurementTests : TestBase
{
    /// <summary>
    ///     Runs an action while sampling the managed heap, and reports the highest reading.
    ///     <para>
    ///         Cumulative allocation is the wrong number here: a renderer allocates and releases
    ///         many times over, so the total says how much work it did, not how much it held at
    ///         once. What the in-memory limit is about is residency, so residency is what is
    ///         sampled.
    ///     </para>
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

        var settled = GC.GetTotalMemory(false) - baseline;
        return Math.Max(peak, settled);
    }

    /// <summary>Builds a Word document with enough content to be worth measuring.</summary>
    /// <param name="paragraphs">How many paragraphs to write.</param>
    /// <returns>The document.</returns>
    private static Document ALargeDocument(int paragraphs)
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);
        var line = new string('x', 512);
        for (var i = 0; i < paragraphs; i++) builder.Writeln(line);

        return document;
    }

    [Fact]
    public void AnInMemoryConversion_ShouldCostSeveralTimesItsOutput_WhichIsWhyTheLimitIsLower()
    {
        // Well inside the in-memory limit, so this measures the shape of the cost rather than the
        // refusal.
        var document = ALargeDocument(2000);

        var bytes = Array.Empty<byte>();
        var peak = MeasurePeakHeap(() =>
        {
            bytes = DocumentConverter.ConvertToBytes(document, DocumentType.Word, "pdf");
        });

        Assert.NotEmpty(bytes);

        // Measured on the pinned build: a 2,000-paragraph document produced a ~186 KB PDF while the
        // managed heap peaked ~32.8 MB above baseline — about 176x the result. The dominant term is
        // the document model, not the output, so the output cap is a reduction rather than a memory
        // bound: a request may hold far more than MaxInMemoryOutputBytes even while respecting it.
        // That is the reason the cap is set well under the on-disk one, and the reason a real
        // memory bound would have to price the source document (a design change, not a constant).
        var multiple = (double)peak / bytes.Length;

        Assert.True(multiple > 10,
            $"A {bytes.Length:N0} byte result peaked at {peak:N0} bytes ({multiple:F1}x). If the "
            + "output ever becomes the dominant term, MaxInMemoryOutputBytes would start to be an "
            + "actual memory bound and the reasoning recorded here should be revisited.");
    }

    [Fact]
    public void TheInMemoryLimit_ShouldBeFarBelowTheOnDiskOne()
    {
        // The property that matters, stated so a future change to either constant has to justify
        // itself: a file is written once and streamed, a response is held whole.
        Assert.True(RenderBudget.MaxInMemoryOutputBytes < RenderBudget.MaxOutputBytes / 16,
            "The in-memory limit exists because the on-disk one is the wrong bound for a response "
            + "that is held in memory; bringing them close together removes the reason for it.");
    }

    [SkippableFact]
    public void NeitherTheOutputCapNorTheInputCap_ShouldBeMistakenForAMemoryBound()
    {
        // Unlicensed, Aspose.Words truncates the document it converts, so a 4,000-paragraph source
        // does not produce more output than a 500-paragraph one and this measures the truncation
        // instead of the conversion. It passed in evaluation runs only because another test class
        // had licensed the process (R13-T01); with that fixed, it says so rather than failing.
        SkipInEvaluationMode(AsposeLibraryType.Words,
            "Evaluation mode truncates the document, so document size stops driving output size");

        // Measured across three sizes of the same document shape: peak managed heap tracks the
        // *decompressed* model, and against the compressed input it ran 2,452x, 2,238x and 2,561x
        // (8.5 KB -> 20.8 MB, 12 KB -> 26.9 MB, 26 KB -> 67.4 MB). A DOCX of repetitive text is the
        // extreme of that, but the shape of the result is the point: a small file can become a very
        // large model.
        //
        // So neither limit the server currently applies is a memory bound. MaxInMemoryOutputBytes
        // bounds what is handed back; SessionConfig.MaxFileSizeMb bounds what is read off disk;
        // the thing that decides peak residency is between them and is bounded by neither. Pricing
        // it means refusing a *loaded* document above some size before converting it in memory,
        // and what that size should be is a product decision, not a constant this test can pick.
        var document = ALargeDocument(500);
        var small = Array.Empty<byte>();
        var smallPeak = MeasurePeakHeap(() =>
        {
            small = DocumentConverter.ConvertToBytes(document, DocumentType.Word, "pdf");
        });

        var larger = ALargeDocument(4000);
        var big = Array.Empty<byte>();
        var bigPeak = MeasurePeakHeap(() =>
        {
            big = DocumentConverter.ConvertToBytes(larger, DocumentType.Word, "pdf");
        });

        Assert.True(big.Length > small.Length, "the larger document should produce more output");
        Assert.True(bigPeak > smallPeak,
            $"Peak heap did not grow with the document ({smallPeak:N0} -> {bigPeak:N0}). If it ever "
            + "stops tracking the model, the reasoning recorded here no longer applies.");

        // And the peak is not explained by the output: it is the model that costs.
        Assert.True(bigPeak > big.Length * 4,
            $"A {big.Length:N0} byte result peaked at {bigPeak:N0} bytes. The output cap can only "
            + "bound the first of those two numbers.");
    }
}
