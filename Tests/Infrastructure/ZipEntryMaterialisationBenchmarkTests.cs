using System.Diagnostics;
using System.IO.Compression;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     R19-RES03, measured rather than judged on shape.
///     <para>
///         Both <see cref="DocumentSizePreflight" /> and <see cref="MhtExternalReferenceScanner" />
///         count entries inside <c>foreach (var entry in archive.Entries)</c>, so the application's
///         own entry cap is applied after <c>Entries</c> has been touched — and touching it makes
///         the runtime read the whole central directory. Statically that is "the cap comes too
///         late"; whether it is a finding depends on what the central directory of a hostile
///         archive actually costs, against the byte cap that gets there first.
///     </para>
///     <para>
///         This is the measurement §32.5 asks for, kept as a test so the answer stays true. It
///         asserts a bound rather than printing a number, because a benchmark nobody fails is a
///         benchmark nobody reads.
///     </para>
/// </summary>
public class ZipEntryMaterialisationBenchmarkTests : TestBase
{
    /// <summary>How many empty entries the adversarial archive holds.</summary>
    /// <remarks>
    ///     Chosen to sit above every application cap in this repository, so the measurement is of
    ///     the runtime's own work rather than of the guard's.
    /// </remarks>
    private const int Entries = 200_000;

    /// <summary>Builds an archive of many empty entries.</summary>
    /// <param name="path">Where to write it.</param>
    /// <returns>Its size on disk.</returns>
    private static long WriteManyEntryArchive(string path)
    {
        using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            for (var i = 0; i < Entries; i++) archive.CreateEntry($"e{i}");
        }

        return new FileInfo(path).Length;
    }

    [Fact]
    public void ACentralDirectoryOfManyEntries_ShouldCostLessThanTheByteCapAllows()
    {
        var path = CreateTestFilePath("many_entries.zip");
        var onDisk = WriteManyEntryArchive(path);

        var before = GC.GetAllocatedBytesForCurrentThread();
        var clock = Stopwatch.StartNew();

        using (var archive = ZipFile.OpenRead(path))
        {
            // The line under measurement. Everything after it is the application's own cap, which
            // is what R19-RES03 observes arrives second.
            _ = archive.Entries.Count;
        }

        clock.Stop();
        var allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        // The amplification factor is what decides whether this is a finding: an archive small
        // enough to pass an upload limit must not cost orders of magnitude more to enumerate.
        var factor = (double)allocated / onDisk;

        Assert.True(factor < 200,
            $"{Entries:N0} entries in {onDisk:N0} bytes on disk allocated {allocated:N0} bytes "
            + $"({factor:F1}x) in {clock.ElapsedMilliseconds:N0} ms — the central directory is a "
            + "cheap enough read that the application cap arriving second is not exploitable. A "
            + "failure here means it has become one, and the count must move before Entries.");

        Assert.True(clock.ElapsedMilliseconds < 10_000,
            $"reading the central directory of {Entries:N0} entries took "
            + $"{clock.ElapsedMilliseconds:N0} ms");
    }

    [Fact]
    public void TheScannersOwnCap_ShouldRefuseSuchAnArchiveAnyway()
    {
        // The control: whatever the enumeration costs, the guard in front of the conversion still
        // refuses the document rather than scanning all of it.
        var path = CreateTestFilePath("many_entries.xps");
        WriteManyEntryArchive(path);

        var failure = Assert.ThrowsAny<ArgumentException>(() =>
            MhtExternalReferenceScanner.EnsureNoRemoteReferences(path, false, [TestDir]));

        // Refused, and the reason is the scanner's own: it will not treat a document it cannot
        // read as one that scanned clean. Which limit stops it first is an implementation detail;
        // that it stops rather than proceeds is the property.
        Assert.False(string.IsNullOrWhiteSpace(failure.Message));
    }
}
