using System.Globalization;
using System.IO.Compression;
using System.Net;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Email;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Refuses an MHT archive that would make the server fetch something from the network.
///     Aspose.Pdf 23.10.0 exposes an external-resource hook on <c>HtmlLoadOptions</c> only, so an
///     MHT input is the one conversion path where the library's own fetching cannot be intercepted.
///     An MHT file exists to be self-contained, so a legitimate one references only the parts it
///     carries; a remote reference is therefore both unusual and the exact thing that turns a
///     conversion into a request issued by the server. This scanner refuses those files before the
///     converter ever opens them.
/// </summary>
public static class MhtExternalReferenceScanner
{
    /// <summary>
    ///     Attribute and CSS constructs that cause a browser or renderer to fetch a resource.
    ///     Matching on the URI rather than the attribute keeps the check independent of which
    ///     attribute carried it.
    /// </summary>
    /// <summary>Largest archive this guard will decode, in bytes.</summary>
    internal const long MaxArchiveBytes = 256L * 1024 * 1024;

    /// <summary>
    ///     Largest encoded archive handed to the MIME parser, in bytes.
    ///     <para>
    ///         The parser reads and decodes the whole file before returning, so a limit applied to
    ///         what it produced was applied after the memory had already been taken (R4-S02). This
    ///         one is checked against the file on disk, before the parser is called. It is well
    ///         below <see cref="MaxArchiveBytes" /> because a saved web page is a document, not an
    ///         archive format anyone streams.
    ///     </para>
    /// </summary>
    internal const long MaxEncodedArchiveBytes = 64L * 1024 * 1024;

    /// <summary>Largest number of text parts scanned from one archive.</summary>
    private const int MaxParts = 1_000;

    /// <summary>Largest total decoded text scanned from one archive, in characters.</summary>
    private const long MaxDecodedChars = 256L * 1024 * 1024;

    /// <summary>Largest single decoded part, in characters.</summary>
    private const long MaxPartChars = 32L * 1024 * 1024;

    /// <summary>
    ///     Largest ratio of decoded to stored bytes accepted from one container entry. A file that
    ///     expands by more than this is a compression bomb rather than a document, and the size of
    ///     the container says nothing about what decoding it would cost.
    /// </summary>
    private const long MaxExpansionRatio = 200;

    /// <summary>How many embedded documents may be decoded before the archive is refused.</summary>
    private const int MaxEmbeddedDocuments = 32;

    /// <summary>The largest embedded document the scan will decode, in characters.</summary>
    private const int MaxEmbeddedDocumentChars = 200_000;

    /// <summary>
    ///     How many <c>data:</c> levels deep the scan follows before refusing the archive. A
    ///     document inside a document inside a document is not a shape any honest saved page has.
    /// </summary>
    private const int MaxEmbeddedDepth = 4;

    /// <summary>How much of the file to inspect for MIME headers.</summary>
    /// <summary>
    ///     How many times a part is decoded before scanning stops. One round covers the ordinary
    ///     encoded reference and the second covers a doubly encoded one; beyond that a parser no
    ///     longer follows either.
    /// </summary>
    private const int DecodeRounds = 2;

    private const int HeaderProbeChars = 8192;

    /// <summary>
    ///     Maximum number of physical MIME lines retained while proving caller provenance.
    ///     This prevents a small newline-only body from expanding into millions of string and
    ///     array objects before the ordinary decoded-content budgets can run.
    /// </summary>
    private const int MaxProvenanceLines = 100_000;

    /// <summary>A marker used only to determine whether this process is in evaluation mode.</summary>
    private const string EvaluationProbeMarker = "aspose-mcp-evaluation-probe";

    /// <summary>
    ///     A resource reference with no scheme of its own, i.e. one resolved against a base.
    ///     <para>
    ///         Measured with a localhost probe against the pinned Aspose.Pdf: an unquoted
    ///         attribute value, a relative <c>url()</c> and a relative <c>@import</c> are all
    ///         fetched, so all three belong here. A relative <c>srcset</c> candidate was not
    ///         fetched by this version - only the accompanying <c>src</c> was - so srcset is not
    ///         matched (R3-S05).
    ///     </para>
    /// </summary>
    private static readonly Regex RelativeReferencePattern = new(
        @"(?i)\b(?:src|href|background|poster)\s*=\s*[""'](?<rel>(?![a-z][a-z0-9+.-]*:)(?!//)(?!#)[^""'>]+)[""']"
        + @"|(?i)\b(?:src|href|background|poster)\s*=\s*(?<rel2>(?![a-z][a-z0-9+.-]*:)(?!//)(?!#)[^""'\s>]+)"
        + @"|(?i)url\(\s*[""']?\s*(?<rel3>(?![a-z][a-z0-9+.-]*:)(?!//)(?!#)[^""')\s]+)"
        + @"|(?i)@import\s+[""'](?<rel4>(?![a-z][a-z0-9+.-]*:)(?!//)(?!#)[^""']+)[""']",
        RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    private static readonly Regex RemoteUriPattern = new(
        @"(?i)\b(?:src|href|background|poster|data|action|formaction|cite|codebase|longdesc|profile|usemap)\s*=\s*[""']?\s*(?<uri>(?:[a-z][a-z0-9+.-]*:)?\/\/[^""'\s>]+)"
        + @"|(?i)url\(\s*[""']?\s*(?<uri2>(?:[a-z][a-z0-9+.-]*:)?\/\/[^""')\s]+)"
        + @"|(?i)@import\s+[""'](?<uri3>(?:[a-z][a-z0-9+.-]*:)?\/\/[^""']+)"
        + @"|(?i)\b(?:src|href|background|poster)\s*=\s*[""']?\s*(?<uri4>file:[^""'\s>]+)"
        + @"|(?i)url\(\s*[""']?\s*(?<uri5>file:[^""')\s]+)"
        + @"|(?i)@import\s+[""'](?<uri6>file:[^""']+)"
        + @"|\]\(\s*(?<uri7>(?:[a-z][a-z0-9+.-]*:)?\/\/[^)\s]+)"
        + @"|\]\(\s*(?<uri8>file:[^)\s]+)",
        RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    /// <summary>
    ///     Schemes that resolve inside the archive itself and can never reach the network or
    ///     the filesystem. <c>file:</c> is deliberately absent: it does reach the filesystem,
    ///     so where it points has to be checked against the allowlist like any other read.
    /// </summary>
    private static readonly string[] LocalSchemes = ["cid:", "data:"];

    /// <summary>
    ///     A CSS escape: a backslash followed by up to six hex digits and an optional space.
    /// </summary>
    private static readonly Regex CssEscape = new(
        @"\\([0-9a-fA-F]{1,6})[ ]?", RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    /// <summary>
    ///     A <c>data:</c> URI carrying a document of its own, in either of the two encodings the
    ///     syntax allows.
    ///     <para>
    ///         The earlier pattern required the media type to be followed immediately by
    ///         <c>;base64,</c>, so <c>data:text/html;charset=utf-8;base64,…</c> matched nothing —
    ///         measured on the pinned build, that spelling fetched (R9-SEC01). Parameters are now
    ///         part of the match, and a payload with no <c>base64</c> parameter is read as the
    ///         percent-encoded form the syntax also permits.
    ///     </para>
    /// </summary>
    private static readonly Regex DataDocument = new(
        @"data:(?<type>text/html|application/xhtml\+xml|image/svg\+xml|text/css)(?<parameters>;[^,]*)?,"
        + @"(?<payload>[^""'\s>)]+)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase, TimeSpan.FromSeconds(5));

    /// <summary>
    ///     Decodes strictly: invalid bytes throw rather than becoming replacement characters, so a
    ///     payload that is not really UTF-8 is refused instead of silently mangled into text with
    ///     no URL in it.
    /// </summary>
    private static readonly Encoding StrictUtf8 = new UTF8Encoding(false, true);

    /// <summary>
    ///     The character sets an embedded active document may declare. Anything else is refused
    ///     rather than guessed at.
    /// </summary>
    private static readonly string[] SupportedCharsets = ["utf-8", "utf8", "us-ascii", "ascii"];

    /// <summary>
    ///     The media types whose content carries references of its own. Everything else in an
    ///     archive — an image, a font, an arbitrary attachment — is bytes the converter will not
    ///     re-parse, so scanning it would cost decoding budget and find nothing.
    /// </summary>
    private static readonly string[] ActiveResourceTypes =
    [
        "text/css", "text/html", "application/xhtml+xml", "image/svg+xml", "text/plain"
    ];

    /// <summary>
    ///     The stream a part's content is read from.
    ///     <para>
    ///         A seam, because the failure that matters here — a part that cannot be read — cannot
    ///         be produced from a fixture otherwise: the parser hands back its own stream and
    ///         nothing in a crafted archive makes reading it throw. What the fail-closed rule is
    ///         worth is exactly what happens on that path, so the path has to be reachable
    ///         (R8-SEC01).
    ///     </para>
    /// </summary>
    internal static Func<AttachmentBase, Stream?> ReadStreamOf { get; set; } =
        part => part.ContentStream;

    /// <summary>
    ///     Throws when <paramref name="path" /> references a resource the server would have to
    ///     fetch. Content is read through <see cref="MailMessage" /> so quoted-printable and
    ///     base64 parts are decoded first; a scan of the raw bytes would miss an encoded reference.
    /// </summary>
    /// <param name="path">Path to the MHT or MHTML file about to be converted.</param>
    /// <param name="allowExternalResources">
    ///     When <c>true</c> the caller has accepted that conversion may issue outbound requests and
    ///     the scan is skipped. Off by default.
    /// </param>
    /// <param name="allowedBasePaths">
    ///     Admin-configured allowlist. A <c>file:</c> reference inside it is a local read the
    ///     caller could have made anyway; one outside it is refused like a remote reference.
    /// </param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the archive references a remote resource and the caller has not opted in.
    /// </exception>
    public static void EnsureSelfContained(string path, bool allowExternalResources = false,
        IReadOnlyList<string>? allowedBasePaths = null)
    {
        if (allowExternalResources) return;

        // Nothing verifiable means nothing to clear. Treating that as "no references found"
        // would wave through exactly the files this guard cannot speak for.
        // Parsed once and reused. Checking readability and then scanning used to load and
        // decode the whole archive twice, before the converter's own parser read it a third
        // time (R2-R04).
        var archive = LooksLikeMimeArchive(path) ? ReadTextParts(path) : new ArchiveContent([], [], false);
        if (archive.Parts.Count == 0)
            throw new ArgumentException(
                "The MHT file could not be read as a MIME archive, so it cannot be checked for "
                + "external references. Conversion was refused. Supply a readable, self-contained "
                + "archive.",
                nameof(path));

        var references = FindArchiveReferences(archive, allowedBasePaths);

        if (references.Count == 0) return;

        throw new ArgumentException(
            "The MHT file references resources that are not contained in the archive, so converting "
            + "it would make the server fetch them. Conversion was refused. Supply a self-contained "
            + "archive, or set allowExternalResources to true to accept the outbound requests. "
            + $"First reference: {Describe(references[0])}");
    }

    /// <summary>
    ///     Throws when a non-archive input carries a reference the server would have to fetch.
    ///     <para>
    ///         Measured against the pinned Aspose.Pdf: converting HTML, Markdown, SVG or EPUB to
    ///         PDF fetches the URLs the document names, and
    ///         <c>HtmlLoadOptions.CustomLoaderOfExternalResources</c> is never consulted for an
    ///         <c>img src</c> on that path - an instrumented strategy recorded zero invocations
    ///         while the probe server still logged the request. The callback therefore protects
    ///         nothing here, and the only thing that does is refusing the input before it is
    ///         opened, exactly as MHT already is.
    ///     </para>
    /// </summary>
    /// <param name="path">Path to the document about to be converted.</param>
    /// <param name="allowExternalResources">
    ///     When <c>true</c> the caller has accepted the outbound requests and the scan is skipped.
    /// </param>
    /// <param name="allowedBasePaths">Roots a local <c>file:</c> reference may point into.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the document references a resource outside itself.
    /// </exception>
    public static void EnsureNoRemoteReferences(string path, bool allowExternalResources = false,
        IReadOnlyList<string>? allowedBasePaths = null)
    {
        if (allowExternalResources) return;

        var parts = ReadConvertibleParts(path);
        if (parts.Count == 0)
            throw new ArgumentException(
                "The document could not be read as text, so it cannot be checked for external "
                + "references. Conversion was refused.", nameof(path));

        var references = ScanParts(parts, allowedBasePaths);
        if (references.Count == 0) return;

        throw new ArgumentException(
            "The document references resources it does not contain, so converting it would make "
            + "the server fetch them. Conversion was refused. Supply a self-contained document, or "
            + "set allowExternalResources to true to accept the outbound requests. "
            + $"First reference: {Describe(references[0])}",
            nameof(path));
    }

    /// <summary>
    ///     Whether a file's own bytes say it is a ZIP container.
    /// </summary>
    /// <param name="path">The file to inspect.</param>
    /// <returns><c>true</c> when the file begins with the ZIP local-header signature.</returns>
    private static bool LooksLikeZipContainer(string path)
    {
        try
        {
            using var file = File.OpenRead(path);
            Span<byte> signature = stackalloc byte[4];
            return file.ReadAtLeast(signature, 4, false) == 4
                   && signature[0] == 0x50 && signature[1] == 0x4B
                   && signature[2] == 0x03 && signature[3] == 0x04;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>
    ///     Reads the scannable text of a convertible input: the file itself for a text format, or
    ///     every text entry for a container format such as EPUB.
    /// </summary>
    /// <param name="path">Path to the document.</param>
    /// <returns>The text parts to scan.</returns>
    private static List<string> ReadConvertibleParts(string path)
    {
        var parts = new List<string>();
        try
        {
            if (new FileInfo(path).Length > MaxArchiveBytes)
                throw new ArgumentException(
                    $"The document is above the {MaxArchiveBytes:N0} byte limit for "
                    + "external-reference scanning.", nameof(path));

            // Any ZIP container, not only the one extension that was named. EPUB was special-cased
            // and everything else fell through to "read the file as one blob of text", so a
            // reference inside a compressed entry of an XPS, an OXPS or any other packaged format
            // was never text this scan could see (R8-SEC02). The container is decided by its bytes,
            // because that is what the converter's own reader goes by.
            if (LooksLikeZipContainer(path))
            {
                // Refused on the count the trailer declares, before the archive is opened.
                // Counting inside the loop below only bounded the loop: `Entries` had already
                // built every entry from the central directory by then (R21-RES01).
                ZipInventoryPreflight.EnsureEntryCountWithin(path, MaxParts, nameof(path));

                using var archive = ZipFile.OpenRead(path);
                var decoded = 0L;
                foreach (var entry in archive.Entries)
                {
                    // Skipping or stopping silently turned a document too large to scan into one
                    // that scanned clean, which is the opposite of what a fail-closed guard owes
                    // its caller (R3-R01). Each of these is a refusal.
                    if (parts.Count >= MaxParts)
                        throw new ArgumentException(
                            $"The document holds more than {MaxParts:N0} parts, which is above the "
                            + "limit for external-reference scanning.", nameof(path));

                    EnsureEntryWithinBudget(entry.Length, entry.CompressedLength, entry.FullName,
                        nameof(path));

                    using var stream = entry.Open();
                    var text = ReadBounded(stream, Math.Min(MaxPartChars, MaxDecodedChars - decoded),
                        entry.FullName, nameof(path));
                    if (text.Length == 0) continue;
                    parts.Add(text);
                    decoded += text.Length;
                }

                return parts;
            }

            using var file = File.OpenRead(path);
            parts.Add(ReadBounded(file, MaxPartChars, Path.GetFileName(path), nameof(path)));
            return parts;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException
                                       or InvalidDataException)
        {
            // Unreadable means unverifiable, and the converter runs its own parser over the same
            // bytes; failing closed is the only answer this guard can give.
            //
            // Returning what had been collected so far was not failing closed: a container whose
            // first entries were harmless and whose later one failed to decode produced a
            // non-empty, clean-looking prefix, which the caller reads as "no remote references"
            // and hands to the converter — which has its own parser and may well get further
            // (R4-S03). A refusal is the only answer that matches what this guard knows.
            throw new ArgumentException(
                $"'{Path.GetFileName(path)}' could not be read in full for external-reference "
                + $"scanning ({ex.GetType().Name}), so it was refused rather than scanned in part.",
                nameof(path));
        }
    }

    /// <summary>
    ///     Refuses a container entry whose decoded size, or the work of decoding it, is above what
    ///     this guard will spend on one part.
    /// </summary>
    /// <param name="decodedBytes">Size the entry decodes to, as the container declares it.</param>
    /// <param name="storedBytes">Size the entry occupies in the container.</param>
    /// <param name="entryName">Entry name, for the error message.</param>
    /// <param name="paramName">Public parameter responsible for the input.</param>
    /// <exception cref="ArgumentException">Thrown when the entry is above either limit.</exception>
    private static void EnsureEntryWithinBudget(long decodedBytes, long storedBytes, string entryName,
        string paramName)
    {
        if (decodedBytes > MaxPartChars)
            throw new ArgumentException(
                $"Part '{entryName}' decodes to {decodedBytes:N0} bytes, above the "
                + $"{MaxPartChars:N0} byte limit for external-reference scanning.", paramName);

        if (storedBytes > 0 && decodedBytes / storedBytes > MaxExpansionRatio)
            throw new ArgumentException(
                $"Part '{entryName}' expands more than {MaxExpansionRatio}-fold when decoded, so "
                + "the document was refused rather than decoded.", paramName);
    }

    /// <summary>
    ///     Reads a stream as text, refusing rather than truncating once the limit is passed.
    /// </summary>
    /// <param name="stream">The stream to read.</param>
    /// <param name="maxChars">Largest number of characters accepted.</param>
    /// <param name="what">Name of what is being read, for the error message.</param>
    /// <param name="paramName">Public parameter responsible for the input.</param>
    /// <returns>The text, which is never longer than <paramref name="maxChars" />.</returns>
    /// <exception cref="ArgumentException">Thrown when the stream holds more than the limit.</exception>
    private static string ReadBounded(Stream stream, long maxChars, string what, string paramName)
    {
        if (maxChars <= 0)
            throw new ArgumentException(
                "The document decodes to more text than the external-reference scan will read, so "
                + "it was refused rather than scanned in part.", paramName);

        using var reader = new StreamReader(stream, leaveOpen: true);
        var builder = new StringBuilder();
        var buffer = new char[81_920];

        while (true)
        {
            var read = reader.Read(buffer, 0, buffer.Length);
            if (read == 0) break;

            // One character past the limit is enough to know: the text is never materialised in
            // full, so the refusal costs no more memory than the limit itself.
            if (builder.Length + read > maxChars)
                throw new ArgumentException(
                    $"'{what}' holds more than {maxChars:N0} characters, above the limit for "
                    + "external-reference scanning.", paramName);

            builder.Append(buffer, 0, read);
        }

        return builder.ToString();
    }

    /// <summary>
    ///     Returns every remote reference found in the archive's decoded text parts.
    /// </summary>
    /// <param name="path">Path to the MHT or MHTML file.</param>
    /// <param name="allowedBasePaths">
    ///     Admin-configured allowlist, used to decide whether a <c>file:</c> reference is a local
    ///     read the caller could have made anyway.
    /// </param>
    /// <returns>The distinct remote URIs, in the order they were found.</returns>
    public static IReadOnlyList<string> FindRemoteReferences(string path,
        IReadOnlyList<string>? allowedBasePaths = null)
    {
        if (!LooksLikeMimeArchive(path)) return [];

        var archive = ReadTextParts(path);
        return FindArchiveReferences(archive, allowedBasePaths);
    }

    /// <summary>Finds decoded references and restores any caller-owned notice URI stripped from vendor output.</summary>
    /// <param name="archive">Decoded archive plus provenance learned from its original bytes.</param>
    /// <param name="allowedBasePaths">Roots a local file reference may use.</param>
    /// <returns>Distinct references in discovery order.</returns>
    private static List<string> FindArchiveReferences(ArchiveContent archive,
        IReadOnlyList<string>? allowedBasePaths)
    {
        var found = ScanParts(archive.Parts.Select(part => part.Text).ToList(),
            allowedBasePaths).ToList();

        // A relative reference is only remote because of where the archive says it came from, so
        // it is resolved against the declared base rather than pattern-matched (R2-S11).
        found.AddRange(FindUnresolvedRelativeReferences(archive));

        // Aspose evaluation mode can collapse a caller-supplied copy of its complete notice into
        // the one decoration it injects. Once decoded, those two origins are indistinguishable.
        // The original active MIME parts are therefore inspected separately; stripping parser
        // output never erases the caller's own notice from the security decision.
        if (archive.CallerCarriesEvaluationNotice
            && !found.Contains(EvaluationNoticeUri, StringComparer.OrdinalIgnoreCase))
            found.Add(EvaluationNoticeUri);

        return found;
    }

    /// <summary>
    ///     Finds relative references that resolve to another host because the archive says so.
    ///     <para>
    ///         A saved web page carries the URL it came from in each part's
    ///         <c>Content-Location</c>, and its markup refers to siblings by relative name. Those
    ///         names normally resolve to other parts of the archive, which is why a relative
    ///         reference is not remote by itself. When the archive holds no part with that name,
    ///         the reference resolves against the remote base instead - and measured against the
    ///         pinned Aspose.Pdf, the converter does exactly that and fetches it. The pattern-based
    ///         scan above cannot see this: the reference carries no scheme and the base lives in a
    ///         MIME header rather than in the body.
    ///     </para>
    /// </summary>
    /// <param name="archive">The archive's parts and the addresses it carries.</param>
    /// <returns>The resolved remote URLs the archive would have to fetch.</returns>
    private static List<string> FindUnresolvedRelativeReferences(ArchiveContent archive)
    {
        List<string> found = [];
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);

        foreach (var part in archive.Parts)
        {
            // A part with no declared address of its own cannot resolve a relative reference, and
            // borrowing another part's address is how one part's origin used to decide where
            // another's references pointed (R8-SEC01).
            if (part.Base == null || part.Base.IsFile) continue;

            foreach (Match match in RelativeReferencePattern.Matches(part.Text))
            {
                var reference = FirstNonEmptyGroup(match, "rel", "rel2", "rel3", "rel4");
                if (reference.Length == 0) continue;
                if (!Uri.TryCreate(part.Base, reference, out var resolved)) continue;
                if (resolved.IsFile) continue;

                // Only the address itself counts as contained. Matching on the file name meant a
                // logo.png the archive carried from one host marked every logo.png as present,
                // whatever host the reference resolved to (R8-SEC01).
                if (archive.Contained.Contains(resolved.AbsoluteUri)) continue;

                if (seen.Add(resolved.AbsoluteUri)) found.Add(resolved.AbsoluteUri);
            }
        }

        return found;
    }

    /// <summary>
    ///     Resolves CSS escapes, so a scheme spelled <c>http\3a //host</c> is seen as the URL the
    ///     style engine will build from it.
    ///     <para>
    ///         Measured against the pinned Aspose build with a loopback listener: that spelling
    ///         produced a real outbound GET while this scanner saw no scheme at all (R8-SEC02).
    ///     </para>
    /// </summary>
    /// <param name="content">The text to unescape.</param>
    /// <returns>The text with CSS escapes resolved.</returns>
    private static string ResolveCssEscapes(string content)
    {
        if (!content.Contains('\\')) return content;

        return CssEscape.Replace(content, match =>
        {
            var value = int.Parse(match.Groups[1].Value, NumberStyles.HexNumber,
                CultureInfo.InvariantCulture);

            return value is > 0 and <= 0x10FFFF and not (>= 0xD800 and <= 0xDFFF)
                ? char.ConvertFromUtf32(value)
                : match.Value;
        });
    }

    /// <summary>
    ///     Reads backslashes as the separators the URL parsers treat them as.
    ///     <para>
    ///         <c>http:\\host\file</c> fetched on the pinned build while the scanner's patterns,
    ///         which look for <c>//</c>, matched nothing (R8-SEC02).
    ///     </para>
    /// </summary>
    /// <param name="content">The text to normalise.</param>
    /// <returns>The text with backslashes read as forward slashes.</returns>
    private static string NormaliseSeparators(string content)
    {
        return content.Contains('\\') ? content.Replace('\\', '/') : content;
    }

    /// <summary>
    ///     The documents carried inside base64 <c>data:</c> URIs.
    ///     <para>
    ///         A <c>data:</c> URI fetches nothing by itself, which is why it is treated as local.
    ///         The document inside one is another matter: an HTML or SVG payload carries its own
    ///         references, and both were fetched on the pinned build while the scan stopped at the
    ///         <c>data:</c> prefix (R8-SEC02).
    ///     </para>
    /// </summary>
    /// <param name="content">The text to look in.</param>
    /// <returns>The decoded payloads, at every level.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the archive carries more embedded documents, a larger one, or a deeper
    ///     nesting than the scan will decode. Each of those used to end the search quietly and
    ///     report the archive clean — measured on the pinned build, a reference in the 33rd
    ///     document, one padded past the size cap, and one two levels down all fetched
    ///     (R9-SEC01). Not scanning something is not the same as finding nothing in it.
    /// </exception>
    private static List<string> EmbeddedDocuments(string content)
    {
        List<string> found = [];
        Queue<string> pending = new();
        pending.Enqueue(content);

        // Widening the readings multiplies how many strings each level produces, and many of them
        // are the same string reached a different way. Searching one twice costs work and finds
        // nothing new, so each is searched once — which is what keeps the closure bounded rather
        // than merely capped (§23.3.3).
        var searched = new HashSet<string>(StringComparer.Ordinal) { content };

        var decoded = 0;
        for (var depth = 0; depth < MaxEmbeddedDepth && pending.Count > 0; depth++)
        for (var remaining = pending.Count; remaining > 0; remaining--)
        {
            var text = pending.Dequeue();
            foreach (var groups in DataDocument.Matches(text).Select(match => match.Groups))
            {
                if (++decoded > MaxEmbeddedDocuments)
                    throw new ArgumentException(
                        $"The document carries more than {MaxEmbeddedDocuments} embedded "
                        + "documents, which is more than can be checked for external "
                        + "references. Conversion was refused.");

                var payload = groups["payload"].Value;
                if (payload.Length > MaxEmbeddedDocumentChars)
                    throw new ArgumentException(
                        "The document carries an embedded document larger than "
                        + $"{MaxEmbeddedDocumentChars:N0} characters, which is more than can "
                        + "be checked for external references. Conversion was refused.");

                var inner = DecodeEmbeddedPayload(payload, groups["parameters"].Value);
                if (inner.Length == 0) continue;

                found.Add(inner);

                // Every reading of the decoded document is searched for the next level, not
                // just its raw text: a `data:` URI that only appears after the inner payload's
                // own CSS escapes or backslashes are resolved was produced by the scanner and
                // then never decoded (§22.4.2).
                foreach (var reading in Readings(inner).Where(searched.Add))
                    pending.Enqueue(reading);
            }
        }

        if (pending.Any(text => DataDocument.IsMatch(text)))
            throw new ArgumentException(
                $"The document nests embedded documents more than {MaxEmbeddedDepth} levels deep, "
                + "which is more than can be checked for external references. Conversion was "
                + "refused.");

        return found;
    }

    /// <summary>
    ///     Decodes one <c>data:</c> payload, in whichever of the two encodings it uses.
    /// </summary>
    /// <param name="payload">The payload as written.</param>
    /// <param name="parameters">The media type's parameters, which say whether it is base64.</param>
    /// <returns>The decoded document.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the payload cannot be decoded. Returning an empty string reported an
    ///     undecodable active document as one holding nothing, which is the same fail-open the
    ///     caps had: the scan cannot say a document is clean when it could not read it (§21.4).
    /// </exception>
    private static string DecodeEmbeddedPayload(string payload, string parameters)
    {
        // A charset the scan does not decode is a document the scan cannot read. Every payload
        // used to be decoded as UTF-8 whatever the media type said, so `charset=utf-16le` came
        // back as UTF-8 nonsense carrying NUL bytes — the scanner produced a string, found no
        // URL in it, and reported the archive clean (§22.4.1).
        var charset = CharsetOf(parameters);
        if (charset.Length > 0 && !SupportedCharsets.Contains(charset, StringComparer.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"The document carries an embedded document declared as '{charset}', which this "
                + "scan does not decode, so it cannot be checked for external references. "
                + "Conversion was refused.");

        try
        {
            return IsBase64(parameters)
                ? StrictUtf8.GetString(Convert.FromBase64String(payload))
                : Uri.UnescapeDataString(payload);
        }
        catch (Exception ex) when (ex is FormatException or UriFormatException
                                       or DecoderFallbackException)
        {
            throw new ArgumentException(
                "The document carries an embedded document that could not be decoded, so it "
                + $"cannot be checked for external references. Conversion was refused. ({ex.Message})");
        }
    }

    /// <summary>Reads the <c>charset</c> parameter of a media type, if it declares one.</summary>
    /// <param name="parameters">The parameter string, starting with its leading semicolon.</param>
    /// <returns>The declared charset, or an empty string.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the media type declares <c>charset</c> more than once, and when it declares
    ///     one with no value. Taking the first of several means decoding by a charset the parser
    ///     may not be using, and a declaration that contradicts itself cannot be decoded with
    ///     confidence — so it is refused rather than guessed at (§23.3.2). An empty declaration is
    ///     the same problem said differently: it returned the same empty string as "no charset
    ///     was declared", which is the one answer that skips the gate entirely and decodes as
    ///     UTF-8 regardless — so a document that named a character set was read as though it had
    ///     named none (R13-SEC02).
    /// </exception>
    private static string CharsetOf(string parameters)
    {
        var declared = new List<string>();

        foreach (var token in parameters.Split(';',
                     StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var separator = token.IndexOf('=');
            if (separator <= 0) continue;

            if (token[..separator].Trim().Equals("charset", StringComparison.OrdinalIgnoreCase))
                declared.Add(token[(separator + 1)..].Trim().Trim('"'));
        }

        if (declared.Count > 1)
            throw new ArgumentException(
                "The document carries an embedded document that declares its character set more "
                + $"than once ({string.Join(", ", declared)}), so it cannot be decoded with "
                + "confidence or checked for external references. Conversion was refused.");

        if (declared is [{ Length: 0 }])
            throw new ArgumentException(
                "The document carries an embedded document that declares a character set without "
                + "naming one, so it cannot be decoded with confidence or checked for external "
                + "references. Conversion was refused.");

        return declared.Count == 1 ? declared[0] : string.Empty;
    }

    /// <summary>
    ///     The readings of one piece of text: as written, with CSS escapes resolved, with
    ///     backslashes read as separators, and with both.
    /// </summary>
    /// <param name="content">The text to read.</param>
    /// <returns>The distinct readings.</returns>
    private static IEnumerable<string> Readings(string content)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);

        // Character references first, then the other three applied to each. Leaving this one out
        // is what made `data&#58;text/html;base64,…` inside a decoded payload invisible: the
        // scanner produced the string and never recognised it as a carrier, and the pinned build
        // fetched what was inside it (§23.3).
        foreach (var decoded in CharacterReferenceVariants(content))
        {
            var unescaped = ResolveCssEscapes(decoded);
            var normalisedRaw = NormaliseSeparators(decoded);
            var normalised = NormaliseSeparators(unescaped);

            foreach (var reading in new[] { decoded, unescaped, normalisedRaw, normalised }.Where(seen.Add))
                yield return reading;
        }
    }

    /// <summary>
    ///     Whether a media type's parameters carry the standalone <c>base64</c> flag.
    ///     <para>
    ///         Searching the whole parameter string for the word matched a parameter merely
    ///         <em>named</em> like it — <c>;charset=base64</c> — and sent a percent-encoded payload
    ///         down the base64 path, where the decode failed and the document was treated as empty
    ///         (§21.4). The flag is a parameter of its own, so it is read as one.
    ///     </para>
    /// </summary>
    /// <param name="parameters">The parameter string, starting with its leading semicolon.</param>
    /// <returns><c>true</c> when base64 is present as a standalone parameter.</returns>
    private static bool IsBase64(string parameters)
    {
        return parameters
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Any(token => string.Equals(token, "base64", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    ///     The text as written, followed by what a parser sees after it resolves character
    ///     references.
    ///     <para>
    ///         Measured against the pinned Aspose build with a loopback listener: an
    ///         <c>img src</c> of <c>http&amp;#x3a;&amp;#x2f;&amp;#x2f;127.0.0.1:PORT/probe.png</c>
    ///         produced a real outbound GET for HTML, SVG and EPUB, while this scanner saw nothing
    ///         — it looks for a literal <c>://</c>, and the encoded form has none until the parser
    ///         decodes it (R7-SEC01). Scanning the decoded form as well closes that gap, and the
    ///         raw form is still scanned because decoding can only remove text, never add it.
    ///     </para>
    ///     <para>
    ///         Decoding repeats a bounded number of times so a doubly encoded reference cannot
    ///         hide one round deeper, and stops as soon as a round changes nothing.
    ///     </para>
    /// </summary>
    /// <param name="content">The part's text.</param>
    /// <returns>The variants to scan, starting with the text as written.</returns>
    private static IEnumerable<string> CharacterReferenceVariants(string content)
    {
        yield return content;

        var current = content;
        for (var round = 0; round < DecodeRounds; round++)
        {
            string decoded;
            try
            {
                decoded = WebUtility.HtmlDecode(current);
            }
            catch (ArgumentException)
            {
                yield break;
            }

            if (string.Equals(decoded, current, StringComparison.Ordinal)) yield break;

            yield return decoded;
            current = decoded;
        }
    }

    /// <summary>
    ///     Every form of one part the scan looks at: as written, as a parser resolves it, and the
    ///     documents it carries inside <c>data:</c> URIs.
    /// </summary>
    /// <param name="content">The part's text.</param>
    /// <returns>The variants to scan.</returns>
    private static IEnumerable<string> ScanVariants(string content)
    {
        foreach (var variant in CharacterReferenceVariants(content))
        {
            yield return variant;

            var unescaped = ResolveCssEscapes(variant);
            if (!string.Equals(unescaped, variant, StringComparison.Ordinal)) yield return unescaped;

            // Both readings of a backslash, because they are exclusive: CSS treats it as an escape
            // introducer and consumes what follows, so `http:\127.0.0.1` came out as one
            // character and the URL disappeared. A URL parser reads the same bytes as separators.
            var normalisedRaw = NormaliseSeparators(variant);
            if (!string.Equals(normalisedRaw, variant, StringComparison.Ordinal))
                yield return normalisedRaw;

            var normalised = NormaliseSeparators(unescaped);
            if (!string.Equals(normalised, unescaped, StringComparison.Ordinal)
                && !string.Equals(normalised, normalisedRaw, StringComparison.Ordinal))
                yield return normalised;

            // Looked for in every reading of the part, not only the raw one: a `data:` URI that
            // exists only after CSS escapes are resolved, or after backslashes are read as
            // separators, was never looked inside (§21.4).
            var carriers = new List<string> { variant };
            if (!string.Equals(unescaped, variant, StringComparison.Ordinal)) carriers.Add(unescaped);
            if (!string.Equals(normalisedRaw, variant, StringComparison.Ordinal))
                carriers.Add(normalisedRaw);
            if (!string.Equals(normalised, unescaped, StringComparison.Ordinal)
                && !string.Equals(normalised, normalisedRaw, StringComparison.Ordinal))
                carriers.Add(normalised);

            // An embedded document gets the same readings as the part that carried it. Applying
            // only character references and CSS escapes inside one left a backslash-written URL in
            // a decoded payload unseen, and it fetched (R9-SEC01).
            foreach (var embedded in carriers.SelectMany(EmbeddedDocuments).Distinct(StringComparer.Ordinal))
            foreach (var inner in CharacterReferenceVariants(embedded))
            {
                yield return inner;

                var innerUnescaped = ResolveCssEscapes(inner);
                if (!string.Equals(innerUnescaped, inner, StringComparison.Ordinal))
                    yield return innerUnescaped;

                var innerNormalisedRaw = NormaliseSeparators(inner);
                if (!string.Equals(innerNormalisedRaw, inner, StringComparison.Ordinal))
                    yield return innerNormalisedRaw;

                var innerNormalised = NormaliseSeparators(innerUnescaped);
                if (!string.Equals(innerNormalised, innerUnescaped, StringComparison.Ordinal)
                    && !string.Equals(innerNormalised, innerNormalisedRaw, StringComparison.Ordinal))
                    yield return innerNormalised;
            }
        }
    }

    /// <summary>
    ///     Scans already-decoded text parts, so an archive is never parsed twice for one decision.
    /// </summary>
    /// <param name="parts">Decoded text parts of the archive.</param>
    /// <param name="allowedBasePaths">Roots the caller may read from.</param>
    /// <returns>The distinct remote URIs, in the order they were found.</returns>
    private static IReadOnlyList<string> ScanParts(List<string> parts,
        IReadOnlyList<string>? allowedBasePaths)
    {
        List<string> found = [];
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);

        foreach (var content in parts)
        foreach (var variant in ScanVariants(content))
        foreach (Match match in RemoteUriPattern.Matches(variant))
        {
            var uri = FirstNonEmpty(match);
            if (uri.Length == 0 || IsLocal(uri)) continue;
            if (IsPermittedFileReference(uri, allowedBasePaths)) continue;
            if (seen.Add(uri)) found.Add(uri);
        }

        return found;
    }

    /// <summary>
    ///     Whether the file's own bytes announce a MIME archive.
    ///     <para>
    ///         Asking the parser is not enough: an unlicensed Aspose.Email adds its evaluation
    ///         notice to the body, so even a file with nothing readable in it comes back
    ///         carrying content. The headers are in the file or they are not, which is true in
    ///         both licensing modes.
    ///     </para>
    /// </summary>
    /// <param name="path">Path to the candidate archive.</param>
    /// <returns><c>true</c> when the leading bytes carry MIME headers.</returns>
    private static bool LooksLikeMimeArchive(string path)
    {
        try
        {
            using var reader = new StreamReader(path);
            var inspected = 0;
            var sawHeader = false;
            var sawMimeHeader = false;

            while (true)
            {
                var line = reader.ReadLine();
                if (line == null) return false;

                inspected += line.Length + 2;
                if (inspected > HeaderProbeChars) return false;
                if (line.Length == 0) return sawHeader && sawMimeHeader;

                if (line[0] is ' ' or '\t')
                {
                    if (!sawHeader) return false;
                    continue;
                }

                var separator = line.IndexOf(':');
                if (separator <= 0 || !line[..separator].All(IsMimeHeaderNameCharacter))
                    return false;

                sawHeader = true;
                var name = line[..separator];
                if (name.Equals("MIME-Version", StringComparison.OrdinalIgnoreCase)
                    || name.Equals("Content-Type", StringComparison.OrdinalIgnoreCase))
                    sawMimeHeader = true;
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>Whether a character may appear in an RFC-style MIME field name.</summary>
    /// <param name="value">The character to inspect.</param>
    /// <returns><c>true</c> when the character is valid in a MIME field name.</returns>
    private static bool IsMimeHeaderNameCharacter(char value)
    {
        return value is >= (char)33 and <= (char)126 && value != ':';
    }

    /// <summary>
    ///     Reads the decoded text of every part that can carry a reference.
    /// </summary>
    /// <param name="path">Path to the MHT or MHTML file.</param>
    /// <returns>
    ///     Decoded HTML and text content of the message and its alternate views. Empty when the
    ///     file yielded nothing readable, which callers must treat as unverifiable rather than
    ///     as clean.
    /// </returns>
    private static ArchiveContent ReadTextParts(string path)
    {
        var parts = new List<ArchivePart>();
        var contained = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var empty = new ArchiveContent([], contained, false);

        // An archive is attacker-supplied input that this guard decodes in full, so the amount
        // of work it can ask for is bounded before any of it is done (R2-R04).
        try
        {
            var length = new FileInfo(path).Length;
            if (length > MaxEncodedArchiveBytes)
                throw new ArgumentException(
                    $"The MHT file is {length:N0} bytes, above the {MaxEncodedArchiveBytes:N0} byte "
                    + "limit for external-reference scanning. The limit applies to the file as it "
                    + "is on disk, because the parser reads all of it before anything can be "
                    + "measured.", nameof(path));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return empty;
        }

        var originalInspected = TryInspectOriginalEvaluationNotice(path,
            out var callerCarriesEvaluationNotice);

        MailMessage message;
        try
        {
            message = MailMessage.Load(path, new MhtmlLoadOptions());
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            // The converter runs a different parser over the same bytes, so "this parser could
            // not read it" says nothing about what that one would fetch. An empty result makes
            // the caller fail closed.
            _ = ex;
            return empty;
        }

        var decodedChars = 0L;
        var stripEvaluationNotice = originalInspected
                                    && message.HtmlBody?.Contains(EvaluationNoticeHtml,
                                        StringComparison.Ordinal) == true
                                    && AsposeAddsEvaluationNotice();

        void Take(string text, Uri? partBase, string paramName)
        {
            if (text.Length == 0) return;

            // Dropping a part once the caps were reached left the rest of the archive unscanned
            // and reported as clean (R3-R01).
            if (text.Length > MaxPartChars || parts.Count >= MaxParts
                                           || decodedChars + text.Length > MaxDecodedChars)
                throw new ArgumentException(
                    "The MHT file decodes to more content than the external-reference scan will "
                    + "read, so it was refused rather than scanned in part.", paramName);

            parts.Add(new ArchivePart(text, partBase));
            decodedChars += text.Length;
        }

        var declaredLocations = new List<string>();

        // The message owns the streams and buffers its parser filled from the file, and it was
        // simply returned from: over repeated attacker-supplied inputs those accumulated until a
        // collection happened to run. Take() throws on any cap, so the release has to be in a
        // finally rather than after the reads (R5-S01).
        using (message)
        {
            // A linked resource is a part the archive carries, so its address is one the archive
            // can satisfy without fetching anything. Its *content* is a different question: a
            // stylesheet, an SVG or an HTML fragment carries references of its own, and treating
            // every linked resource as nothing but a name left those unscanned (R9-SEC02).
            foreach (var resource in message.LinkedResources)
            {
                var resourceBase = LocationOf(resource, contained, declaredLocations);
                if (CarriesReferences(resource))
                    Take(ReadPart(resource, "a linked resource", nameof(path)), resourceBase,
                        nameof(path));
            }

            // The bodies have no headers of their own; each alternate view does, and carries the
            // same markup. A body is therefore scanned for absolute references but cannot resolve
            // a relative one, which is the view's job.
            if (!string.IsNullOrEmpty(message.HtmlBody))
                Take(RemoveInjectedEvaluationNotice(message.HtmlBody, stripEvaluationNotice), null,
                    nameof(path));
            if (!string.IsNullOrEmpty(message.Body))
                Take(RemoveInjectedEvaluationNotice(message.Body, stripEvaluationNotice), null,
                    nameof(path));

            foreach (var view in message.AlternateViews)
                Take(RemoveInjectedEvaluationNotice(ReadPart(view, "an alternate view", nameof(path)),
                        stripEvaluationNotice),
                    LocationOf(view, contained, declaredLocations), nameof(path));
        }

        // A part may declare its address relatively, as a saved page's image usually does. Those
        // are resolved against the archive's own primary address so the archive is credited with
        // carrying them — which is safe in a way sharing a base for *references* is not: it can
        // only ever name a part the archive actually holds, never decide where a fetch would go.
        var primaryBase = parts.Select(part => part.Base).FirstOrDefault(uri => uri != null);
        if (primaryBase != null)
            foreach (var declared in declaredLocations)
                if (!Uri.TryCreate(declared, UriKind.Absolute, out _)
                    && Uri.TryCreate(primaryBase, declared, out var resolved))
                    contained.Add(resolved.AbsoluteUri);

        return new ArchiveContent(parts, contained, callerCarriesEvaluationNotice);
    }

    /// <summary>Records and resolves the address a MIME part declares for itself.</summary>
    /// <param name="attachment">The parsed MIME part.</param>
    /// <param name="contained">Addresses the archive can satisfy internally.</param>
    /// <param name="declaredLocations">Raw declared locations, including relative ones.</param>
    /// <returns>The absolute location, or null when none was declared or it is relative.</returns>
    private static Uri? LocationOf(AttachmentBase attachment, HashSet<string> contained,
        List<string> declaredLocations)
    {
        var declared = attachment.Headers["Content-Location"]?.Trim();
        if (string.IsNullOrEmpty(declared)) return null;

        contained.Add(declared);
        declaredLocations.Add(declared);
        if (!Uri.TryCreate(declared, UriKind.Absolute, out var absolute)) return null;

        contained.Add(absolute.AbsoluteUri);
        return absolute;
    }

    /// <summary>
    ///     Whether a part's declared media type is one whose content can name another resource.
    /// </summary>
    /// <param name="part">The archive part.</param>
    /// <returns><c>true</c> when the part's content is worth scanning.</returns>
    private static bool CarriesReferences(AttachmentBase part)
    {
        var type = part.ContentType?.MediaType ?? string.Empty;
        return ActiveResourceTypes.Contains(type, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    ///     Reads one part's content stream as text.
    /// </summary>
    /// <param name="part">The archive part to read.</param>
    /// <param name="what">What the part is, for the size message.</param>
    /// <param name="paramName">Public parameter responsible for the input.</param>
    /// <returns>The part's text.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the part is longer than one part may be, or cannot be read at all. A part
    ///     that could not be read used to come back as an empty string, so the archive was scanned
    ///     without it and reported clean — the one outcome a guard like this may never produce
    ///     (R8-SEC01), and a part with no stream at all went the same way (R9-SEC02).
    /// </exception>
    private static string ReadPart(AttachmentBase part, string what, string paramName)
    {
        try
        {
            // A part with no stream is one that was not read, not one that held nothing.
            var stream = ReadStreamOf(part)
                         ?? throw new ArgumentException(
                             "A part of the MHT file has no readable content, so the archive "
                             + "cannot be checked for external references. Conversion was "
                             + "refused.");

            var position = stream.CanSeek ? stream.Position : 0;
            if (stream.CanSeek) stream.Position = 0;
            // Bounded like every other decoded part: ReadToEnd here had no limit of its own, so
            // one oversized view was read in full whatever the aggregate cap said (R4-S02).
            var text = ReadBounded(stream, MaxPartChars, what, paramName);
            if (stream.CanSeek) stream.Position = position;
            return text;
        }
        catch (Exception ex) when (ex is IOException or ObjectDisposedException or NotSupportedException)
        {
            throw new ArgumentException(
                "A part of the MHT file could not be read, so the archive cannot be checked for "
                + $"external references. Conversion was refused. ({ex.Message})");
        }
    }

    /// <summary>Returns the first of the named groups that captured something.</summary>
    /// <param name="match">The match to read.</param>
    /// <param name="names">Group names, in priority order.</param>
    /// <returns>The captured value, trimmed, or an empty string.</returns>
    private static string FirstNonEmptyGroup(Match match, params string[] names)
    {
        foreach (var name in names)
        {
            var group = match.Groups[name];
            if (group is { Success: true, Value.Length: > 0 }) return group.Value.Trim();
        }

        return string.Empty;
    }

    /// <summary>
    ///     Returns whichever named group of the alternation actually matched.
    /// </summary>
    /// <param name="match">A successful match of <see cref="RemoteUriPattern" />.</param>
    /// <returns>The captured URI, or an empty string.</returns>
    private static string FirstNonEmpty(Match match)
    {
        foreach (var name in new[] { "uri", "uri2", "uri3", "uri4", "uri5", "uri6", "uri7", "uri8" })
        {
            var group = match.Groups[name];
            if (group is { Success: true, Value.Length: > 0 }) return group.Value.Trim();
        }

        return string.Empty;
    }

    /// <summary>
    ///     Whether a URI resolves inside the archive or the local machine rather than the network.
    /// </summary>
    /// <param name="uri">The captured URI.</param>
    /// <returns><c>true</c> when no outbound request would be made.</returns>
    private static bool IsLocal(string uri)
    {
        return LocalSchemes.Any(scheme => uri.StartsWith(scheme, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    ///     Whether a <c>file:</c> reference points somewhere the caller is already allowed to
    ///     read. Anything else — including any <c>file:</c> URI when no allowlist is configured —
    ///     stays a reference the conversion must not follow on the caller's behalf.
    /// </summary>
    /// <param name="uri">The captured URI.</param>
    /// <param name="allowedBasePaths">Roots the caller may read from.</param>
    /// <returns><c>true</c> when the reference is a permitted local file.</returns>
    private static bool IsPermittedFileReference(string uri, IReadOnlyList<string>? allowedBasePaths)
    {
        if (!uri.StartsWith("file:", StringComparison.OrdinalIgnoreCase)) return false;
        if (allowedBasePaths == null || allowedBasePaths.Count == 0) return false;

        if (!Uri.TryCreate(uri, UriKind.Absolute, out var parsed) || !parsed.IsFile) return false;

        try
        {
            SecurityHelper.ResolveAndEnsureWithinAllowlist(parsed.LocalPath, allowedBasePaths, "reference");
            return true;
        }
        catch (Exception ex) when (ex is ArgumentException or IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>
    ///     Probes the library itself to distinguish its evaluation decoration from identical
    ///     caller input. A licensed process adds nothing; an unlicensed process adds the known
    ///     fragment before the marker.
    /// </summary>
    /// <returns><c>true</c> only when the current library injects the measured fragment.</returns>
    private static bool AsposeAddsEvaluationNotice()
    {
        var archive = string.Join("\r\n",
            "From: <aspose-mcp-probe>",
            "Subject: probe",
            "MIME-Version: 1.0",
            "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "",
            $"<html><body>{EvaluationProbeMarker}</body></html>",
            "");

        try
        {
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(archive));
            using var message = MailMessage.Load(stream, new MhtmlLoadOptions());
            var html = message.HtmlBody ?? string.Empty;
            var notice = html.IndexOf(EvaluationNoticeHtml, StringComparison.Ordinal);
            var marker = html.IndexOf(EvaluationProbeMarker, StringComparison.Ordinal);
            return notice >= 0 && marker > notice;
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            // Failure to prove the injection means nothing is removed. The resulting vendor URL
            // is then treated as remote, which refuses conversion rather than weakening the guard.
            return false;
        }
    }

    /// <summary>
    ///     Inspects caller-owned MIME bytes independently of Aspose's evaluation decoration.
    /// </summary>
    /// <param name="path">Archive path already checked against the encoded-size limit.</param>
    /// <param name="carriesNotice">
    ///     Whether a literal, base64, or quoted-printable active caller body contains the complete
    ///     notice markup.
    /// </param>
    /// <returns>
    ///     <c>true</c> when every declared transfer encoding was understood and the inspection is
    ///     authoritative; <c>false</c> when vendor output must not be stripped.
    /// </returns>
    private static bool TryInspectOriginalEvaluationNotice(string path, out bool carriesNotice)
    {
        carriesNotice = false;

        try
        {
            var bytes = File.ReadAllBytes(path);
            if (!TryReadProvenanceLines(bytes, out var lines)
                || !TryDescribeOriginalTextParts(lines, out var parts))
                return false;

            foreach (var part in parts)
            {
                var bodyEnd = part.BodyStart;
                while (bodyEnd < lines.Count
                       && (part.EndBoundary == null
                           || !IsBoundaryDelimiterLine(lines[bodyEnd], part.EndBoundary, out _)))
                    bodyEnd++;

                if (part.EndBoundary != null && bodyEnd == lines.Count) return false;

                byte[] decoded;
                if (part.TransferEncoding.Equals("base64", StringComparison.OrdinalIgnoreCase))
                {
                    var payload = string.Concat(lines
                        .Skip(part.BodyStart)
                        .Take(bodyEnd - part.BodyStart)
                        .Select(line => line.Trim()));
                    decoded = Convert.FromBase64String(payload);
                }
                else if (part.TransferEncoding.Equals("quoted-printable",
                             StringComparison.OrdinalIgnoreCase))
                {
                    decoded = DecodeQuotedPrintable(string.Join("\n", lines
                        .Skip(part.BodyStart)
                        .Take(bodyEnd - part.BodyStart)));
                }
                else if (part.TransferEncoding.Equals("7bit", StringComparison.OrdinalIgnoreCase)
                         || part.TransferEncoding.Equals("8bit", StringComparison.OrdinalIgnoreCase)
                         || part.TransferEncoding.Equals("binary", StringComparison.OrdinalIgnoreCase))
                {
                    decoded = Encoding.Latin1.GetBytes(string.Join("\n", lines
                        .Skip(part.BodyStart)
                        .Take(bodyEnd - part.BodyStart)));
                }
                else
                {
                    // Unknown caller encoding means origin cannot be proved. The caller remains
                    // safe because the evaluation decoration will not be stripped.
                    return false;
                }

                if (!TryGetStrictEncoding(part.Charset, out var charset)) return false;

                string text;
                try
                {
                    text = charset.GetString(decoded);
                }
                catch (DecoderFallbackException)
                {
                    return false;
                }

                if (!text.Contains(EvaluationNoticeHtml, StringComparison.Ordinal)) continue;
                carriesNotice = true;
                return true;
            }

            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or FormatException)
        {
            // Failure to inspect original bytes is not evidence that the notice is vendor-owned.
            // Returning false disables stripping, so the remote URI is refused instead.
            return false;
        }
    }

    /// <summary>Reads MIME lines without constructing whole-file replacement strings.</summary>
    /// <param name="bytes">The original MHT bytes.</param>
    /// <param name="lines">Receives the MIME lines when the bounded read succeeds.</param>
    /// <returns><c>true</c> when the complete input fits within the provenance line budget.</returns>
    private static bool TryReadProvenanceLines(byte[] bytes, out List<string> lines)
    {
        lines = [];
        using var stream = new MemoryStream(bytes, false);
        using var reader = new StreamReader(stream, Encoding.Latin1, false, 4096, true);

        while (reader.ReadLine() is { } line)
        {
            if (lines.Count >= MaxProvenanceLines)
            {
                lines = [];
                return false;
            }

            if (lines.Count == 0 && line.StartsWith("\u00EF\u00BB\u00BF", StringComparison.Ordinal))
                line = line[3..];

            lines.Add(line);
        }

        return true;
    }

    /// <summary>
    ///     Builds a conservative MIME part model. A malformed header block is not guessed at:
    ///     failing this proof simply keeps the vendor notice in the security scan.
    /// </summary>
    /// <param name="lines">The original MIME lines to describe.</param>
    /// <param name="parts">Receives the active text parts when the MIME structure is valid.</param>
    /// <returns><c>true</c> when the MIME structure was described completely and unambiguously.</returns>
    private static bool TryDescribeOriginalTextParts(IReadOnlyList<string> lines,
        out List<OriginalTextPart> parts)
    {
        parts = [];
        var boundaries = new HashSet<string>(StringComparer.Ordinal);
        var inHeaders = true;
        string? openingBoundary = null;
        StringBuilder? contentType = null;
        StringBuilder? transferEncoding = null;
        string? lastHeader = null;

        for (var index = 0; index < lines.Count; index++)
        {
            var line = lines[index];
            if (!inHeaders)
            {
                var marker = boundaries.FirstOrDefault(boundary =>
                    IsBoundaryDelimiterLine(line, boundary, out _));
                if (marker == null) continue;

                _ = IsBoundaryDelimiterLine(line, marker, out var closing);
                if (closing)
                {
                    openingBoundary = null;
                    continue;
                }

                inHeaders = true;
                openingBoundary = marker;
                contentType = null;
                transferEncoding = null;
                lastHeader = null;
                continue;
            }

            if (line.Length == 0)
            {
                if (!TryAddOriginalTextPart(contentType?.ToString(), transferEncoding?.ToString(), openingBoundary,
                        index + 1, boundaries, parts))
                    return false;

                inHeaders = false;
                continue;
            }

            if (line[0] is ' ' or '\t')
            {
                if (lastHeader == null) return false;
                if (lastHeader.Equals("Content-Type", StringComparison.OrdinalIgnoreCase))
                    contentType?.Append(' ').Append(line.Trim());
                else if (lastHeader.Equals("Content-Transfer-Encoding",
                             StringComparison.OrdinalIgnoreCase))
                    transferEncoding?.Append(' ').Append(line.Trim());
                continue;
            }

            var separator = line.IndexOf(':');
            if (separator <= 0 || !line[..separator].All(IsMimeHeaderNameCharacter)) return false;

            lastHeader = line[..separator];
            var value = line[(separator + 1)..].Trim();
            if (lastHeader.Equals("Content-Type", StringComparison.OrdinalIgnoreCase))
            {
                if (contentType != null) return false;
                contentType = new StringBuilder(value);
            }
            else if (lastHeader.Equals("Content-Transfer-Encoding",
                         StringComparison.OrdinalIgnoreCase))
            {
                if (transferEncoding != null) return false;
                transferEncoding = new StringBuilder(value);
            }
        }

        return !inHeaders && parts.Count > 0;
    }

    /// <summary>Matches one declared MIME boundary with optional transport-added padding.</summary>
    /// <param name="line">The MIME line to inspect.</param>
    /// <param name="boundary">The declared boundary, including its leading dashes.</param>
    /// <param name="closing">Receives whether the line is the closing form of the boundary.</param>
    /// <returns><c>true</c> when the line contains exactly the declared boundary and allowed padding.</returns>
    private static bool IsBoundaryDelimiterLine(string line, string boundary, out bool closing)
    {
        var closeMarker = boundary + "--";
        if (line.StartsWith(closeMarker, StringComparison.Ordinal)
            && ContainsOnlyTransportPadding(line, closeMarker.Length))
        {
            closing = true;
            return true;
        }

        closing = false;
        return line.StartsWith(boundary, StringComparison.Ordinal)
               && ContainsOnlyTransportPadding(line, boundary.Length);
    }

    /// <summary>Checks the optional space or tab padding allowed after a MIME delimiter.</summary>
    /// <param name="line">The delimiter line to inspect.</param>
    /// <param name="start">The first character after the declared delimiter.</param>
    /// <returns><c>true</c> when every remaining character is transport padding.</returns>
    private static bool ContainsOnlyTransportPadding(string line, int start)
    {
        for (var index = start; index < line.Length; index++)
            if (line[index] is not (' ' or '\t'))
                return false;

        return true;
    }

    /// <summary>Completes one MIME header block and records active text content.</summary>
    /// <param name="contentType">The unfolded Content-Type header, when declared.</param>
    /// <param name="transferEncoding">The unfolded Content-Transfer-Encoding header, when declared.</param>
    /// <param name="openingBoundary">The boundary that opened this part, when nested.</param>
    /// <param name="bodyStart">The index of the first body line.</param>
    /// <param name="boundaries">Known multipart boundaries, updated for a multipart part.</param>
    /// <param name="parts">The collection that receives active text parts.</param>
    /// <returns>
    ///     <c>true</c> when the header block is supported and any active text part was recorded;
    ///     supported inactive parts are accepted and skipped.
    /// </returns>
    private static bool TryAddOriginalTextPart(string? contentType, string? transferEncoding,
        string? openingBoundary, int bodyStart, HashSet<string> boundaries,
        List<OriginalTextPart> parts)
    {
        var mediaType = "text/plain";
        var charset = string.Empty;

        if (contentType != null)
        {
            ContentType parsed;
            try
            {
                parsed = new ContentType(contentType);
            }
            catch (FormatException)
            {
                return false;
            }

            mediaType = parsed.MediaType;
            charset = parsed.CharSet ?? string.Empty;
            if (mediaType.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrWhiteSpace(parsed.Boundary)
                    || parsed.Boundary.Length > HeaderProbeChars)
                    return false;

                boundaries.Add("--" + parsed.Boundary);
                return transferEncoding == null
                       || transferEncoding.Equals("7bit", StringComparison.OrdinalIgnoreCase)
                       || transferEncoding.Equals("8bit", StringComparison.OrdinalIgnoreCase)
                       || transferEncoding.Equals("binary", StringComparison.OrdinalIgnoreCase);
            }
        }

        if (!ActiveResourceTypes.Contains(mediaType, StringComparer.OrdinalIgnoreCase)) return true;

        parts.Add(new OriginalTextPart(bodyStart, openingBoundary,
            string.IsNullOrWhiteSpace(transferEncoding) ? "7bit" : transferEncoding.Trim(), charset));
        return true;
    }

    /// <summary>Resolves a declared MIME charset with strict decoding semantics.</summary>
    /// <param name="charset">The declared charset, or an empty value for strict US-ASCII.</param>
    /// <param name="encoding">Receives the strict encoding when resolution succeeds.</param>
    /// <returns><c>true</c> when the charset is known and can be decoded without replacement.</returns>
    private static bool TryGetStrictEncoding(string charset, out Encoding encoding)
    {
        try
        {
            encoding = string.IsNullOrWhiteSpace(charset)
                ? Encoding.GetEncoding("us-ascii", EncoderFallback.ExceptionFallback,
                    DecoderFallback.ExceptionFallback)
                : Encoding.GetEncoding(charset.Trim(), EncoderFallback.ExceptionFallback,
                    DecoderFallback.ExceptionFallback);
            return true;
        }
        catch (ArgumentException)
        {
            encoding = Encoding.ASCII;
            return false;
        }
    }

    /// <summary>Decodes the transfer-encoding features needed to recognise caller provenance.</summary>
    /// <param name="value">The quoted-printable body text.</param>
    /// <returns>The decoded body bytes.</returns>
    private static byte[] DecodeQuotedPrintable(string value)
    {
        using var decoded = new MemoryStream(value.Length);

        for (var index = 0; index < value.Length; index++)
        {
            if (value[index] == '=' && index + 1 < value.Length && value[index + 1] == '\n')
            {
                index++;
                continue;
            }

            if (value[index] == '=' && index + 2 < value.Length
                                    && byte.TryParse(value.AsSpan(index + 1, 2),
                                        NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var escaped))
            {
                decoded.WriteByte(escaped);
                index += 2;
                continue;
            }

            decoded.WriteByte((byte)value[index]);
        }

        return decoded.ToArray();
    }

    /// <summary>Removes one proven library-injected decoration, never a URI by value.</summary>
    /// <param name="content">Decoded HTML produced by Aspose.Email.</param>
    /// <param name="strip">Whether a same-process probe observed the injection.</param>
    /// <returns>The caller's content with one leading library decoration removed.</returns>
    private static string RemoveInjectedEvaluationNotice(string content, bool strip)
    {
        if (!strip) return content;

        var index = content.IndexOf(EvaluationNoticeHtml, StringComparison.Ordinal);
        return index < 0 ? content : content.Remove(index, EvaluationNoticeHtml.Length);
    }

    /// <summary>
    ///     Shortens a URI for an error message so a long query string cannot flood the response.
    /// </summary>
    /// <param name="uri">The URI to describe.</param>
    /// <returns>The URI, truncated when long.</returns>
    private static string Describe(string uri)
    {
        return uri.Length <= 120 ? uri : uri[..120] + "...";
    }

    /// <summary>
    ///     One decoded part of an archive, with the address the archive says it came from.
    /// </summary>
    /// <param name="Text">The part's decoded text.</param>
    /// <param name="Base">
    ///     The part's own <c>Content-Location</c>, taken from the headers the MIME parser read.
    ///     A relative reference in this part resolves against this and nothing else: every part
    ///     used to share the first remote location found anywhere in the file, so one part's
    ///     origin decided where another part's references pointed (R8-SEC01).
    /// </param>
    private sealed record ArchivePart(string Text, Uri? Base);

    /// <summary>
    ///     An archive's decoded parts and the addresses it can satisfy from inside itself.
    /// </summary>
    /// <param name="Parts">The decoded text parts.</param>
    /// <param name="Contained">
    ///     Every address the archive carries a part for, as the parser reported it. These used to
    ///     be scraped out of the raw file with a regex, so a <c>Content-Location:</c> line written
    ///     inside an HTML body was read as if it were a part header and could mark any address as
    ///     contained (R8-SEC01).
    /// </param>
    /// <param name="CallerCarriesEvaluationNotice">
    ///     Whether an original active caller part, independently decoded, contained the complete
    ///     vendor notice before Aspose could add its own decoration.
    /// </param>
    private sealed record ArchiveContent(
        List<ArchivePart> Parts,
        HashSet<string> Contained,
        bool CallerCarriesEvaluationNotice);

    /// <summary>Caller-owned active text part used only for evaluation-notice provenance.</summary>
    private sealed record OriginalTextPart(
        int BodyStart,
        string? EndBoundary,
        string TransferEncoding,
        string Charset);

    /// <summary>
    ///     The markup Aspose.Email 23.10 prepends to HTML bodies in evaluation mode. The URL is
    ///     comparison data, not a destination this server opens. The complete injected fragment
    ///     is removed only after a same-process probe proves the library is currently adding it;
    ///     matching the URL alone let caller-supplied markup opt itself out of the guard.
    /// </summary>
#pragma warning disable S5332 // The HTTP URL is vendor-generated comparison data, never requested here.
#pragma warning disable S1075 // The complete vendor fragment must be matched exactly to remove only its injection.
    private const string EvaluationNoticeUri =
        "http://www.aspose.com/corporate/purchase/end-user-license-agreement.aspx";

    private const string EvaluationNoticeHtml =
        "<br><center><span style=\"color:red\">Evaluation Only. Created with Aspose.Email for .NET. "
        + "Copyright 2002-2022 Aspose Pty Ltd.</span></center><br><center><a href=\""
        + EvaluationNoticeUri + "\">"
        + "View EULA Online</a></center><hr><br>";
#pragma warning restore S1075
#pragma warning restore S5332
}
