using System.Reflection;
using System.Text;
using Aspose.Words;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Core.Conversion;

/// <summary>
///     What an in-memory conversion is allowed to produce, and what the extension bridge is allowed
///     to keep.
///     <para>
///         Neither had a byte limit. The conversion wrote into a plain <see cref="MemoryStream" />,
///         and <c>PixelBudget</c> — the only limit on the path — bounds the raster dimensions a
///         render is asked for, which says nothing about how many bytes a compressed image or a
///         non-image format then writes. The bridge then kept the result, bounding how many
///         conversions it held and not how large any of them was, so a full cache of maximum-sized
///         conversions was tens of gigabytes (R4-R01).
///     </para>
/// </summary>
public class ConversionOutputBudgetTests
{
    private static readonly long MaxCachedConversionBytes =
        (long)typeof(ExtensionSessionBridge)
            .GetField("MaxCachedConversionBytes", BindingFlags.NonPublic | BindingFlags.Static)!
            .GetValue(null)!;

    private static readonly long MaxTotalCachedBytes =
        (long)typeof(ExtensionSessionBridge)
            .GetField("MaxTotalCachedBytes", BindingFlags.NonPublic | BindingFlags.Static)!
            .GetValue(null)!;

    /// <summary>
    ///     Every way this codebase has ever written a conversion result to disk.
    ///     <para>
    ///         The guard used to look for one spelling — <c>X.Save(resolved…)</c> — which is how
    ///         nine other sinks stayed invisible: a <c>File.WriteAllText</c> for PDF→TXT, four
    ///         <c>new FileStream(…, FileMode.Create)</c> devices, two <c>SheetRender.ToImage</c>
    ///         path overloads and two more in the PDF image handler. They truncated the caller's
    ///         destination before the output was known to fit and left a partial file when it did
    ///         not (R8-C02). Naming the kinds is what makes the next one visible.
    ///     </para>
    /// </summary>
    private static readonly string[] RawDiskSinks =
    [
        "File.WriteAllText(", "File.WriteAllBytes(", "File.WriteAllLines(", "File.AppendAllText(",
        "File.Create(", "File.OpenWrite(", "new FileStream(", "new StreamWriter(",
        ".Save(resolved", ".ToImage(0, resolved", ".Process(resolved"
    ];

    /// <summary>
    ///     The source files this guard reads. It is the conversion surface and nothing else: the
    ///     converter itself, and the PDF image extraction that shares its fan-out.
    ///     <para>
    ///         The rest of the server writes to disk too, and this guard has never had anything to
    ///         say about it — the name it used to carry, "every disk conversion sink", read as
    ///         though it did (R9-C01). Those writes are covered by
    ///         <c>FanOutBudgetInventoryTests</c>, which requires a ceiling on any handler writing
    ///         into a caller-named directory, and are otherwise single files whose size is the
    ///         size the caller asked for.
    ///     </para>
    /// </summary>
    private static readonly string[] ConversionSources =
    [
        "Core/Conversion/DocumentConverter.cs",
        "Handlers/Pdf/Image/ExtractPdfImageHandler.cs"
    ];

    /// <summary>
    ///     Builds a document with enough content that any format writes well over a few hundred bytes.
    /// </summary>
    /// <returns>The document.</returns>
    private static Document DocumentWithContent()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < 200; i++)
            builder.Writeln($"Line {i} of a document that has to be worth more than a few hundred bytes.");

        return doc;
    }

    /// <summary>
    ///     Both an image format and a non-image one: the limit belongs to the conversion, not to the
    ///     raster path that happened to have one.
    /// </summary>
    /// <param name="format">The output format to convert to.</param>
    [Theory]
    [InlineData("pdf")]
    [InlineData("png")]
    public void AConversionOverItsByteLimit_ShouldBeRefusedWithoutReturningPartialBytes(string format)
    {
        var document = DocumentWithContent();

        var refusal = Assert.Throws<ArgumentException>(() =>
            DocumentConverter.ConvertToStream(document, DocumentType.Word, format, 512));

        Assert.Contains("conversion output", refusal.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("pdf")]
    [InlineData("png")]
    public void AConversionInsideItsByteLimit_ShouldStillBeProduced(string format)
    {
        var document = DocumentWithContent();

        using var stream = DocumentConverter.ConvertToStream(document, DocumentType.Word, format,
            64L * 1024 * 1024);

        Assert.True(stream.Length > 512, "the fixture converted nothing, so the limit proved nothing");
    }

    /// <summary>
    ///     A conversion larger than the per-entry limit is still handed back — it is a valid
    ///     conversion — but it is not kept, because keeping it is what made the cache unbounded.
    /// </summary>
    [Fact]
    public void TheCache_ShouldAdmitByBytesRatherThanOnlyByCount()
    {
        var mayBeCached = typeof(ExtensionSessionBridge)
            .GetMethod("MayBeCached", BindingFlags.NonPublic | BindingFlags.Static)!;

        Assert.True((bool)mayBeCached.Invoke(null, [new byte[MaxCachedConversionBytes]])!);
        Assert.False((bool)mayBeCached.Invoke(null, [new byte[MaxCachedConversionBytes + 1]])!);
    }

    /// <summary>
    ///     §17.4.2: a result held in memory is bounded far below one written to a file.
    ///     <para>
    ///         The on-disk limit is right for a file, which is written once and streamed. An
    ///         in-memory response is held at least twice over at the moment it is copied out, and
    ///         concurrent sessions multiply that, so the same two-gigabyte bound meant a single
    ///         request could hold several gigabytes of managed memory.
    ///     </para>
    /// </summary>
    [Fact]
    public void TheInMemoryLimit_ShouldBeFarBelowTheOnDiskOne()
    {
        Assert.True(RenderBudget.MaxInMemoryOutputBytes > 0);
        Assert.True(
            RenderBudget.MaxInMemoryOutputBytes
            <= RenderBudget.MaxOutputBytes / 8,
            "an in-memory limit close to the on-disk one bounds nothing that matters");

        // A conversion the in-memory path would refuse is still available as a file.
        Assert.True(MaxCachedConversionBytes
                    <= RenderBudget.MaxInMemoryOutputBytes,
            "the cache must not admit results the in-memory path would not even produce");
    }

    /// <summary>
    ///     The per-entry limit alone does not bound the cache: entries just under it still add up,
    ///     which is why there is a total as well and why it has to be the smaller of the two bounds.
    /// </summary>
    [Fact]
    public void TheCacheLimits_ShouldBoundTheTotalAndNotOnlyEachEntry()
    {
        Assert.True(MaxCachedConversionBytes > 0);
        Assert.True(MaxTotalCachedBytes >= MaxCachedConversionBytes,
            "a total below the per-entry limit would refuse every entry");
        Assert.True(MaxTotalCachedBytes < RenderBudget.MaxOutputBytes,
            "the cache total has to be smaller than a single maximum-sized conversion, or it " +
            "bounds nothing that matters");
    }

    /// <summary>Locates the repository root from the test binary's directory.</summary>
    /// <returns>The repository root.</returns>
    private static DirectoryInfo RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "AsposeMcpServer.csproj")))
            directory = directory.Parent;

        Assert.NotNull(directory);
        return directory;
    }

    /// <summary>
    ///     §18.5.1 / R8-C02: the on-disk limit was a declaration and nothing else — no disk sink
    ///     was attached to it, so the 2 GiB figure described a limit that did not exist. Every
    ///     conversion that writes a file now goes through the same transactional publisher, which
    ///     is what applies the bound and what leaves the destination alone on a refusal.
    ///     <para>
    ///         Scoped to <see cref="ConversionSources" />. It says nothing about the rest of the
    ///         server, and its name no longer suggests otherwise (R9-C01).
    ///     </para>
    /// </summary>
    [Fact]
    public void TheConversionSourcesThisGuardReads_ShouldWriteThroughTheBoundedPublisher()
    {
        var root = RepositoryRoot();
        var offenders = new List<string>();
        var publishes = 0;

        foreach (var relative in ConversionSources)
        {
            var source = File.ReadAllText(
                Path.Combine(root.FullName, relative.Replace('/', Path.DirectorySeparatorChar)),
                Encoding.UTF8);

            publishes += source.Split("BoundedFilePublisher.Publish(").Length - 1;
            offenders.AddRange(RawDiskSinks
                .Where(sink => source.Contains(sink, StringComparison.Ordinal))
                .Select(sink => $"{relative}: {sink}"));
        }

        Assert.True(publishes > 0,
            "No bounded publish was found at all, so this guard is reading the wrong files.");
        Assert.True(offenders.Count == 0,
            "these conversions write to disk without passing through the byte cap:"
            + Environment.NewLine + string.Join(Environment.NewLine, offenders));
    }

    /// <summary>
    ///     Whether a source file puts bytes on disk at a path, by any route — including the
    ///     bounded publisher.
    /// </summary>
    /// <remarks>
    ///     <see cref="RawDiskSinks" /> lists the <em>unmanaged</em> routes, so a file that had
    ///     been fixed to publish through the bound matched none of them. Asking that list whether
    ///     a file writes to disk therefore answered "no" for exactly the files this guard exists
    ///     to cover, and the coverage check passed with the converter removed from it.
    /// </remarks>
    /// <param name="source">The file's text.</param>
    /// <returns><c>true</c> when the file writes to a path.</returns>
    private static bool WritesToDisk(string source)
    {
        return source.Contains("BoundedFilePublisher.Publish(", StringComparison.Ordinal)
               || source.Contains("BoundedFileBatch", StringComparison.Ordinal)
               || RawDiskSinks.Any(sink => source.Contains(sink, StringComparison.Ordinal));
    }

    [Fact]
    public void EveryConversionSourceThatWritesToDisk_ShouldBeListed()
    {
        // The scope above is honest only while it still covers the whole conversion surface. A new
        // file under Core/Conversion writing to disk without being listed is exactly how the
        // previous guard came to describe less than its name claimed (R9-C01).
        var root = RepositoryRoot();
        var conversion = Path.Combine(root.FullName, "Core", "Conversion");

        var unlisted = Directory.EnumerateFiles(conversion, "*.cs", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(root.FullName, path).Replace('\\', '/'))
            .Where(relative => !ConversionSources.Contains(relative, StringComparer.Ordinal))
            .Where(relative => WritesToDisk(File.ReadAllText(Path.Combine(root.FullName,
                relative.Replace('/', Path.DirectorySeparatorChar)))))
            .ToList();

        Assert.True(unlisted.Count == 0,
            "these conversion sources write to disk but are not covered by this guard: "
            + string.Join(", ", unlisted));
    }

    /// <summary>Every sink kind, one case each.</summary>
    /// <returns>The sink patterns as theory data.</returns>
    public static TheoryData<string> SinkKinds()
    {
        var data = new TheoryData<string>();
        foreach (var sink in RawDiskSinks) data.Add(sink);
        return data;
    }

    /// <param name="sink">The sink spelling under test.</param>
    [Theory]
    [MemberData(nameof(SinkKinds))]
    public void EverySinkKind_ShouldBeSeenIfItComesBack(string sink)
    {
        // A pattern that no longer matches anything would sit in the list looking like coverage,
        // so each one is shown a source that uses it.
        var reintroduced = $"        {sink}outputPath);";

        Assert.Contains(sink, reintroduced, StringComparison.Ordinal);
    }

    /// <summary>
    ///     §18.5.1: a conversion that writes a file leaves nothing half-written behind.
    /// </summary>
    [Fact]
    public void ADiskConversion_ShouldLeaveNoStagingFileBehind()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ConvertBudget_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);

        try
        {
            var output = Path.Combine(directory, "converted.pdf");
            DocumentConverter.ConvertWordDocument(DocumentWithContent(), output, "pdf");

            Assert.True(File.Exists(output));
            Assert.DoesNotContain(Directory.GetFiles(directory),
                f => Path.GetFileName(f).Contains(".partial-", StringComparison.Ordinal));
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }
}
