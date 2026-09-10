using System.Buffers.Binary;
using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     R21-RES01 and R21-RES02: an archive is refused on the count its trailer declares, before
///     any entry object exists.
/// </summary>
public class ZipInventoryPreflightTests : IDisposable
{
    private readonly string _directory =
        Path.Combine(Path.GetTempPath(), "ZipPreflight_" + Guid.NewGuid().ToString("N"));

    /// <summary>Creates the working directory.</summary>
    public ZipInventoryPreflightTests()
    {
        Directory.CreateDirectory(_directory);
    }

    /// <summary>Removes it.</summary>
    public void Dispose()
    {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        GC.SuppressFinalize(this);
    }

    /// <summary>Writes an archive of many empty entries.</summary>
    /// <param name="entries">How many.</param>
    /// <returns>Its path.</returns>
    private string ArchiveWith(int entries)
    {
        var path = Path.Combine(_directory, $"{entries}.zip");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        for (var i = 0; i < entries; i++)
            archive.CreateEntry($"e{i}", CompressionLevel.NoCompression);

        return path;
    }

    [Fact]
    public void TheDeclaredCount_ShouldBeReadFromTheTrailer()
    {
        using var stream = File.OpenRead(ArchiveWith(37));
        Assert.Equal(37, ZipInventoryPreflight.EntryCountOf(stream));
    }

    [Fact]
    public void AnArchiveOverTheLimit_ShouldBeRefusedFasterThanItCouldBeOpened()
    {
        // Twenty thousand empty entries is a few hundred kilobytes of central directory. The
        // point is not the absolute time but that refusal costs the trailer, not the directory.
        var path = ArchiveWith(20_000);

        var clock = Stopwatch.StartNew();
        var refusal = Assert.Throws<ArgumentException>(() =>
            ZipInventoryPreflight.EnsureEntryCountWithin(path, 1_000, "path"));
        clock.Stop();

        Assert.Contains("20,000", refusal.Message);
        Assert.True(clock.ElapsedMilliseconds < 500,
            $"refusing on the trailer took {clock.ElapsedMilliseconds} ms");
    }

    [Fact]
    public void AnArchiveWithinTheLimit_ShouldPass()
    {
        var exception = Record.Exception(() =>
            ZipInventoryPreflight.EnsureEntryCountWithin(ArchiveWith(10), 1_000, "path"));

        Assert.Null(exception);
    }

    [Fact]
    public void ATrailerThatCannotBeFound_ShouldBeRefused()
    {
        var path = Path.Combine(_directory, "junk.zip");
        File.WriteAllBytes(path, new byte[64]);

        Assert.Throws<ArgumentException>(() =>
            ZipInventoryPreflight.EnsureEntryCountWithin(path, 1_000, "path"));
    }

    [Fact]
    public void ATrailingRecordInsideAComment_IsTheRecordTheReaderWouldUse()
    {
        // Two self-consistent records: the real one, whose comment holds a second one that ends
        // exactly at the file's end. ZipArchive scans backwards and takes the last, so the
        // preflight must say what the reader would say — the planted count — and refuse on it.
        // Preferring the "intended" record would let a small declared count front a huge one.
        var path = ArchiveWith(3);
        var bytes = File.ReadAllBytes(path).ToList();

        // Append a comment holding a fake record with an absurd count: patch the real record's
        // comment length, then append the fake.
        var fake = new byte[22];
        BinaryPrimitives.WriteUInt32LittleEndian(fake, 0x06054B50);
        BinaryPrimitives.WriteUInt16LittleEndian(fake.AsSpan(8), 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(fake.AsSpan(10), 0xFFFE);
        var recordAt = bytes.Count - 22;
        BinaryPrimitives.WriteUInt16LittleEndian(
            CollectionsMarshal.AsSpan(bytes)[(recordAt + 20)..],
            (ushort)fake.Length);
        bytes.AddRange(fake);
        File.WriteAllBytes(path, bytes.ToArray());

        using var stream = File.OpenRead(path);
        Assert.Equal(0xFFFE, ZipInventoryPreflight.EntryCountOf(stream));
    }

    [Fact]
    public void AZip64Trailer_ShouldBeFollowedToItsRecord()
    {
        // Built by hand: a ZIP64 end-of-central-directory record, its locator, and a 16-bit record
        // whose counts say "look at ZIP64".
        var record64 = new byte[56];
        BinaryPrimitives.WriteUInt32LittleEndian(record64, 0x06064B50);
        BinaryPrimitives.WriteUInt64LittleEndian(record64.AsSpan(24), 123_456);
        BinaryPrimitives.WriteUInt64LittleEndian(record64.AsSpan(32), 123_456);

        var locator = new byte[20];
        BinaryPrimitives.WriteUInt32LittleEndian(locator, 0x07064B50);
        BinaryPrimitives.WriteUInt64LittleEndian(locator.AsSpan(8), 0);

        var record = new byte[22];
        BinaryPrimitives.WriteUInt32LittleEndian(record, 0x06054B50);
        BinaryPrimitives.WriteUInt16LittleEndian(record.AsSpan(8), 0xFFFF);
        BinaryPrimitives.WriteUInt16LittleEndian(record.AsSpan(10), 0xFFFF);

        var path = Path.Combine(_directory, "zip64.zip");
        File.WriteAllBytes(path, record64.Concat(locator).Concat(record).ToArray());

        using var stream = File.OpenRead(path);
        Assert.Equal(123_456, ZipInventoryPreflight.EntryCountOf(stream));
    }

    [Fact]
    public void TheExternalReferenceScanner_ShouldRefuseOnTheDeclaredCountBeforeOpening()
    {
        // R21-RES01 at the caller. The in-loop counter refused after `Entries` had built every
        // part; the preflight refuses on the trailer, and says "declares", not "holds".
        // The ZIP branch lives in the check every non-MHT convertible format goes through; the
        // container is decided by its bytes, not its extension.
        var path = ArchiveWith(20_000);

        var clock = Stopwatch.StartNew();
        var refusal = Assert.Throws<ArgumentException>(() =>
            MhtExternalReferenceScanner.EnsureNoRemoteReferences(path, false, [_directory]));
        clock.Stop();

        Assert.Contains("declares", refusal.Message, StringComparison.Ordinal);
        Assert.True(clock.ElapsedMilliseconds < 1_000,
            $"the scanner took {clock.ElapsedMilliseconds} ms to refuse on the trailer");
    }

    [Fact]
    public void ThePresentationPreflight_ShouldGiveUpOnTheDeclaredCountBeforeOpening()
    {
        // R21-RES02 at the caller. 100,001 empty entries is a few megabytes of central directory
        // the old code materialised in full before its counter said no.
        var path = Path.ChangeExtension(ArchiveWith(100_001), ".pptx");
        File.Move(Path.Combine(_directory, "100001.zip"), path);

        // Outcome and time do not tell the versions apart: the old code also returned null, and
        // it built a hundred thousand entry objects in well under a second. What it could not do
        // is give up without allocating them — so allocation, per thread, is the measurement.
        var before = GC.GetAllocatedBytesForCurrentThread();
        var slides = DocumentSizePreflight.SlideCount(path);
        var allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.Null(slides);
        Assert.True(allocated < 2L * 1024 * 1024,
            $"the preflight allocated {allocated:N0} bytes giving up on an archive it should have "
            + "judged from its trailer");
    }

    [Fact]
    public void AKnownTooLargePresentation_ShouldBeRefusedBeforeAnyLoader_EvenWithNoSlideLimit()
    {
        // R22-RES01. The preflight declined to enumerate 100,001 parts — and then reported
        // "unknown", which the caller read as "no opinion" and passed to the vendor loader.
        // Known too large is a refusal, and it does not wait for a configured slide limit.
        var path = Path.ChangeExtension(ArchiveWith(100_001), ".pptx");
        File.Move(Path.Combine(_directory, "100001.zip"), path);

        var preflight = DocumentSizePreflight.Presentation(path);
        Assert.Null(preflight.Slides);
        Assert.NotNull(preflight.Refusal);

        var refusal = Assert.Throws<ArgumentException>(() =>
            DocumentConverter.EnsureFileWithinLimit(path, DocumentType.PowerPoint, new InMemoryModelLimits()));
        Assert.Contains("100,001", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnUnreadablePresentation_ShouldStillBeUnknown_NotRefused()
    {
        // The control: a file the preflight cannot read is an unknown, and unknown is not refused
        // here — the loaded-model check downstream remains the bound for it.
        var path = Path.Combine(_directory, "not-a-zip.pptx");
        File.WriteAllText(path, "not an archive at all");

        var preflight = DocumentSizePreflight.Presentation(path);
        Assert.Null(preflight.Slides);
        Assert.Null(preflight.Refusal);
        DocumentConverter.EnsureFileWithinLimit(path, DocumentType.PowerPoint, new InMemoryModelLimits());
    }
}
