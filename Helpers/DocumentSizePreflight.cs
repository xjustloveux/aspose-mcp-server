using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Counts what a document holds by reading the file, before any loader opens it.
///     <para>
///         <c>DocumentConverter.EnsureModelWithinLimit</c> measures a model that is already in
///         memory, so it bounds whether to <em>continue</em> — not what the load itself costs. That
///         limitation was recorded for several rounds as needing metadata Aspose does not expose
///         (§23.13.1), which was the wrong conclusion: Aspose is not the only thing that can read
///         these formats. OOXML is a ZIP whose parts are named by what they contain, and a PDF's
///         page tree states its own count. Both are cheap to read and neither builds a model.
///     </para>
///     <para>
///         The PDF side is a scan for the page tree, not a PDF parser: it reads a window at each
///         end of the file and looks for the root page-tree node there. It searched only the tail
///         at first, on the assumption that the catalogue is written last — which is false for the
///         producer this server itself uses. An Aspose-saved document puts its page tree at byte
///         157 and, once an evaluation watermark takes it past 64 KB, that tree fell outside the
///         only window being read, so every such file answered "unknown" and the PDF limit never
///         applied (R13-C01). The fixture that was meant to cover this stayed under the window and
///         so never met the case.
///     </para>
///     <para>
///         Deliberately conservative. When the shape is not one this can read — a legacy binary
///         format, an encrypted package, a PDF whose count is only derivable by walking the tree —
///         it answers "unknown" and the loaded-model check downstream remains the only bound.
///         Refusing on a number this could not establish would turn a cheap optimisation into a
///         source of false refusals.
///     </para>
/// </summary>
public static class DocumentSizePreflight
{
    /// <summary>Largest entry name table this will read, so a crafted archive cannot exhaust memory.</summary>
    private const int MaxEntriesInspected = 100_000;

    /// <summary>How much of a PDF is searched at each end for the page-tree count.</summary>
    /// <remarks>
    ///     Wide enough to clear the <c>/Kids</c> array of a large document, which sits between the
    ///     node's <c>/Type</c> and its <c>/Count</c> and grows with the page count.
    /// </remarks>
    private const int PdfWindowBytes = 256 * 1024;

    /// <summary>Slide parts of a PowerPoint package, which are one per slide.</summary>
    private static readonly Regex SlidePart = new(
        @"^ppt/slides/slide\d+\.xml$", RegexOptions.Compiled | RegexOptions.IgnoreCase,
        TimeSpan.FromSeconds(5));

    /// <summary>Marks a page-tree node, whether the root of the tree or an interior one.</summary>
    private static readonly Regex PdfPagesNode = new(
        @"/Type\s*/Pages\b", RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    /// <summary>A page-tree node's own count.</summary>
    private static readonly Regex PdfCount = new(
        @"/Count\s+(?<count>\d+)", RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    /// <summary>
    ///     How many slides a presentation file holds, without opening it as a presentation.
    /// </summary>
    /// <param name="path">The file to inspect.</param>
    /// <returns>The slide count, or null when it cannot be read this way.</returns>
    public static int? SlideCount(string path)
    {
        return Presentation(path).Slides;
    }

    /// <summary>
    ///     What can be known about a presentation file without opening it as a presentation.
    /// </summary>
    /// <param name="path">The file to inspect.</param>
    /// <returns>The slide count when readable, and a refusal when the file declares itself too large.</returns>
    public static PresentationPreflight Presentation(string path)
    {
        try
        {
            // The count the trailer declares, before anything is opened: the loop's own counter
            // bounded the loop and not the central directory the archive had already built
            // (R21-RES02). Over the limit is a refusal in its own right, not an unknown: the
            // loader must not be the one to find out (R22-RES01).
            using (var trailer = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (ZipInventoryPreflight.EntryCountOf(trailer) is not { } declared)
                    return new PresentationPreflight(null, null);
                if (declared > MaxEntriesInspected)
                    return new PresentationPreflight(null,
                        $"The presentation declares {declared:N0} package parts; at most "
                        + $"{MaxEntriesInspected:N0} are accepted.");
            }

            using var archive = ZipFile.OpenRead(path);

            var slides = 0;
            var inspected = 0;
            foreach (var entry in archive.Entries)
            {
                if (++inspected > MaxEntriesInspected)
                    return new PresentationPreflight(null,
                        $"The presentation holds more than {MaxEntriesInspected:N0} package parts.");
                if (SlidePart.IsMatch(entry.FullName)) slides++;
            }

            // A .ppt (legacy binary) is not a ZIP and never reaches here; a .pptx with no slide
            // parts is not a shape this understands, so it says so rather than reporting zero.
            return new PresentationPreflight(slides > 0 ? slides : null, null);
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException
                                       or UnauthorizedAccessException or NotSupportedException)
        {
            return new PresentationPreflight(null, null);
        }
    }

    /// <summary>
    ///     How many pages a PDF holds, read from the page tree rather than by loading it.
    /// </summary>
    /// <param name="path">The file to inspect.</param>
    /// <returns>The page count, or null when it cannot be read this way.</returns>
    public static int? PageCount(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (stream.Length == 0) return null;

            int? counted = null;
            foreach (var window in Ends(stream))
            foreach (var count in RootPageTreeCounts(window))
            {
                // Two roots disagreeing means the file carries more than one revision of its page
                // tree and this cannot tell which is current. Answering the larger would refuse a
                // document that is within the limit, so it answers nothing.
                if (counted is not null && counted != count) return null;
                counted = count;
            }

            return counted;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or RegexMatchTimeoutException)
        {
            return null;
        }
    }

    /// <summary>Reads a window at each end of the file, without overlapping them.</summary>
    /// <param name="stream">The PDF to read.</param>
    /// <returns>The head window, then the tail window when the file is long enough to have one.</returns>
    private static IEnumerable<string> Ends(FileStream stream)
    {
        // Both ends, because either can hold the catalogue: producers commonly write it first
        // (Aspose puts it at byte 157), and an incremental update appends a newer one last.
        yield return Window(stream, 0, (int)Math.Min(stream.Length, PdfWindowBytes));

        if (stream.Length <= PdfWindowBytes) yield break;

        var tail = (int)Math.Min(stream.Length - PdfWindowBytes, PdfWindowBytes);
        yield return Window(stream, stream.Length - tail, tail);
    }

    /// <summary>Reads one window as text, byte for byte.</summary>
    /// <param name="stream">The PDF to read.</param>
    /// <param name="offset">Where the window starts.</param>
    /// <param name="length">How long the window is.</param>
    /// <returns>The window's bytes as Latin-1 text, which preserves every byte value.</returns>
    private static string Window(FileStream stream, long offset, int length)
    {
        stream.Seek(offset, SeekOrigin.Begin);
        var buffer = new byte[length];
        var read = stream.ReadAtLeast(buffer, length, false);
        return Encoding.Latin1.GetString(buffer, 0, read);
    }

    /// <summary>The counts stated by whichever root page-tree nodes appear in this window.</summary>
    /// <param name="text">One window of the file.</param>
    /// <returns>A count per root node found, which is normally one and occasionally none.</returns>
    private static IEnumerable<int> RootPageTreeCounts(string text)
    {
        foreach (var nodeIndex in PdfPagesNode.Matches(text).Select(node => node.Index))
        {
            // The node's own dictionary, which is flat: /Type, /Kids, /Count, and on every node
            // but the root a /Parent.
            var open = text.LastIndexOf("<<", nodeIndex, StringComparison.Ordinal);
            var close = text.IndexOf(">>", nodeIndex, StringComparison.Ordinal);
            if (open < 0 || close < 0) continue;

            var dictionary = text[open..close];

            // An interior node states how many pages hang below it, not how many the document has.
            // Only the root has no parent, and only the root's count is the document's.
            if (dictionary.Contains("/Parent", StringComparison.Ordinal)) continue;

            var count = PdfCount.Match(dictionary);
            if (count.Success && int.TryParse(count.Groups["count"].Value, out var pages) && pages > 0)
                yield return pages;
        }
    }

    /// <summary>What a preflight found out about a presentation.</summary>
    /// <param name="Slides">The slide count, or null when it could not be read this way.</param>
    /// <param name="Refusal">
    ///     Why the file must not be opened at all, when the preflight established that much: the
    ///     archive declares more parts than a presentation may have. Null otherwise.
    /// </param>
    /// <remarks>
    ///     "Unknown" and "known too large" used to be one null, and the caller treated null as
    ///     "no opinion" and went on to the loader — so the very archive the trailer check refused
    ///     to enumerate was handed to the vendor parser to enumerate (R22-RES01).
    /// </remarks>
    public readonly record struct PresentationPreflight(int? Slides, string? Refusal);
}
