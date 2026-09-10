using System.Buffers.Binary;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Answers "how many entries does this archive claim" without opening the archive.
///     <para>
///         Both readers that count entries — the external-reference scanner and the presentation
///         preflight — did so by enumerating <c>ZipArchive.Entries</c>, which builds every entry
///         object from the central directory before the first one is handed back. An archive of
///         a few hundred kilobytes describing a hundred thousand empty parts was therefore fully
///         materialised before the count ever refused it (R21-RES01, R21-RES02).
///     </para>
///     <para>
///         The count is in the end-of-central-directory record, twenty-two bytes at the end of
///         the file plus an optional comment, or in the ZIP64 record that record points to. That
///         is all this reads. An archive whose trailer cannot be found or is inconsistent is
///         refused rather than guessed at: the reader that follows would only be more confused.
///     </para>
/// </summary>
public static class ZipInventoryPreflight
{
    /// <summary>The most bytes the trailer may occupy: the record and the largest comment.</summary>
    private const int MaxTrailerBytes = 22 + 0xFFFF;

    private const uint EndOfCentralDirectory = 0x06054B50;
    private const uint Zip64Locator = 0x07064B50;
    private const uint Zip64EndOfCentralDirectory = 0x06064B50;

    /// <summary>Refuses an archive whose central directory claims more entries than allowed.</summary>
    /// <param name="path">The archive.</param>
    /// <param name="maximumEntries">The most entries a reader may be asked to materialise.</param>
    /// <param name="paramName">What the path is, for the refusal.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the archive claims more entries than allowed, or when its trailer cannot be
    ///     read consistently.
    /// </exception>
    /// <exception cref="IOException">Thrown when the file cannot be read.</exception>
    public static void EnsureEntryCountWithin(string path, long maximumEntries, string paramName)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);

        var claimed = EntryCountOf(stream)
                      ?? throw new ArgumentException(
                          "The archive's central directory record could not be read, so how many "
                          + "parts it holds cannot be known before opening it.", paramName);

        if (claimed > maximumEntries)
            throw new ArgumentException(
                $"The archive declares {claimed:N0} parts, which is above the limit of "
                + $"{maximumEntries:N0} this server will read.", paramName);
    }

    /// <summary>The entry count an archive's trailer declares, or null when it cannot be read.</summary>
    /// <param name="stream">The archive, positioned anywhere; seekable.</param>
    /// <returns>The declared count, or null.</returns>
    /// <remarks>
    ///     Reads the last <see cref="MaxTrailerBytes" /> bytes at most. The end-of-central-directory
    ///     signature is searched backwards, because a comment may follow it; the first match from
    ///     the end whose declared comment length reaches exactly the end of the file is the record.
    ///     A 16-bit count of <c>0xFFFF</c> means ZIP64, whose locator sits twenty bytes before the
    ///     record and names the 64-bit record's offset.
    /// </remarks>
    public static long? EntryCountOf(Stream stream)
    {
        if (!stream.CanSeek || stream.Length < 22) return null;

        var tail = (int)Math.Min(MaxTrailerBytes, stream.Length);
        var buffer = new byte[tail];
        stream.Position = stream.Length - tail;
        ReadExactly(stream, buffer);

        return EntryCountOfTrailer(buffer, stream);
    }

    /// <summary>Parses the count out of a trailer already in memory.</summary>
    /// <param name="tail">The last bytes of the archive.</param>
    /// <param name="whole">The archive, for the ZIP64 record when the trailer points to one; may be null.</param>
    /// <returns>The declared count, or null.</returns>
    internal static long? EntryCountOfTrailer(ReadOnlySpan<byte> tail, Stream? whole)
    {
        for (var i = tail.Length - 22; i >= 0; i--)
        {
            if (BinaryPrimitives.ReadUInt32LittleEndian(tail[i..]) != EndOfCentralDirectory) continue;

            // The record's own comment length must place the end of the comment exactly at the
            // end of the file; otherwise this is a signature inside data, not the record.
            var commentLength = BinaryPrimitives.ReadUInt16LittleEndian(tail[(i + 20)..]);
            if (i + 22 + commentLength != tail.Length) continue;

            var count = BinaryPrimitives.ReadUInt16LittleEndian(tail[(i + 10)..]);
            var total = BinaryPrimitives.ReadUInt16LittleEndian(tail[(i + 8)..]);
            if (count != 0xFFFF && total != 0xFFFF) return Math.Max(count, total);

            // ZIP64: the locator precedes the record by twenty bytes.
            if (i < 20 || whole == null) return null;
            if (BinaryPrimitives.ReadUInt32LittleEndian(tail[(i - 20)..]) != Zip64Locator) return null;

            var recordOffset = BinaryPrimitives.ReadUInt64LittleEndian(tail[(i - 20 + 8)..]);
            if (recordOffset > (ulong)(whole.Length - 56)) return null;

            var record = new byte[56];
            whole.Position = (long)recordOffset;
            ReadExactly(whole, record);
            if (BinaryPrimitives.ReadUInt32LittleEndian(record) != Zip64EndOfCentralDirectory) return null;

            var count64 = BinaryPrimitives.ReadUInt64LittleEndian(record.AsSpan(32));
            return count64 > long.MaxValue ? long.MaxValue : (long)count64;
        }

        return null;
    }

    /// <summary>Fills a buffer or gives up.</summary>
    /// <param name="stream">The source.</param>
    /// <param name="buffer">What to fill.</param>
    /// <exception cref="EndOfStreamException">Thrown when the stream ends first.</exception>
    private static void ReadExactly(Stream stream, byte[] buffer)
    {
        var read = 0;
        while (read < buffer.Length)
        {
            var step = stream.Read(buffer, read, buffer.Length - read);
            if (step <= 0) throw new EndOfStreamException();
            read += step;
        }
    }
}
