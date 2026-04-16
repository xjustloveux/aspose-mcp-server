using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Helpers.Ole;

/// <summary>
///     Projects an Aspose.Words <see cref="Shape" /> carrying an <see cref="OleFormat" />
///     into the cross-tool <see cref="OleObjectMetadata" /> shape. Filename derivation
///     flows through <see cref="OleSanitizerHelper.SanitizeOleFileName" /> so Word /
///     Excel / PowerPoint outputs stay identical for the same raw input (AC-18).
/// </summary>
public static class WordOleMetadataMapper
{
    /// <summary>
    ///     Maps a single Word OLE-bearing shape into the cross-tool metadata record.
    /// </summary>
    /// <param name="shape">
    ///     The shape carrying a non-null <see cref="OleFormat" />. Must not be null.
    /// </param>
    /// <param name="index">Zero-based flat index of this OLE within the document.</param>
    /// <param name="location">Container-specific location (already resolved by the handler).</param>
    /// <param name="sizeBytes">
    ///     Payload size in bytes, typically obtained by saving the OLE payload to a
    ///     <see cref="MemoryStream" />. Zero for linked objects.
    /// </param>
    /// <returns>A fully-populated <see cref="OleObjectMetadata" /> record.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="shape" /> is null.</exception>
    public static OleObjectMetadata Map(Shape shape, int index, OleShapeLocation? location, long sizeBytes)
    {
        ArgumentNullException.ThrowIfNull(shape);
        var ole = shape.OleFormat!;
        var progId = ole.ProgId;
        var rawName = ResolveRawFileName(ole);
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName(rawName, index, progId);

        var extension = OleSanitizerHelper.NormalizeExtension(ole.SuggestedExtension);
        if (string.IsNullOrEmpty(extension))
            extension = OleSanitizerHelper.ExtensionFromProgId(progId);

        string? linkTarget = null;
        if (ole.IsLink)
        {
            var raw = ole.SourceFullName;
            linkTarget = string.IsNullOrEmpty(raw)
                ? null
                : OleSanitizerHelper.SanitizeForLog(raw);
        }

        return new OleObjectMetadata
        {
            Index = index,
            RawFileName = rawName,
            SuggestedFileName = suggested,
            ProgId = progId,
            SizeBytes = sizeBytes,
            IsLinked = ole.IsLink,
            Extension = extension,
            ShapeLocation = location,
            LinkTarget = linkTarget
        };
    }

    /// <summary>
    ///     Selects the best raw-filename candidate off a Word <see cref="OleFormat" />:
    ///     <c>OlePackage.FileName</c> → <c>OlePackage.DisplayName</c> →
    ///     <c>SuggestedFileName</c> → <c>Path.GetFileName(SourceFullName)</c>.
    /// </summary>
    /// <param name="ole">The source OLE format. Must not be null.</param>
    /// <returns>
    ///     The first non-empty candidate, or <c>null</c> when none of the four sources
    ///     yield a non-empty string.
    /// </returns>
    private static string? ResolveRawFileName(OleFormat ole)
    {
        if (ole.OlePackage != null)
        {
            if (!string.IsNullOrEmpty(ole.OlePackage.FileName)) return ole.OlePackage.FileName;
            if (!string.IsNullOrEmpty(ole.OlePackage.DisplayName)) return ole.OlePackage.DisplayName;
        }

        if (!string.IsNullOrEmpty(ole.SuggestedFileName)) return ole.SuggestedFileName;

        if (!string.IsNullOrEmpty(ole.SourceFullName))
            return Path.GetFileName(ole.SourceFullName);

        return null;
    }

    /// <summary>
    ///     Computes the payload size of a Word OLE shape by saving the embedded data to
    ///     an in-memory stream. Returns zero for linked objects (no embedded payload).
    /// </summary>
    /// <param name="shape">OLE-bearing shape. Must not be null.</param>
    /// <returns>Payload size in bytes, or zero for linked OLE objects.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="shape" /> is null.</exception>
    public static long ComputeSize(Shape shape)
    {
        ArgumentNullException.ThrowIfNull(shape);
        if (shape.OleFormat == null || shape.OleFormat.IsLink) return 0;
        using var ms = new MemoryStream();
        shape.OleFormat.Save(ms);
        return ms.Length;
    }

    /// <summary>
    ///     Resolves the location of an OLE shape in Word: the zero-based index of its
    ///     owning section plus the zero-based paragraph index within that section's body.
    /// </summary>
    /// <param name="document">The owning document. Must not be null.</param>
    /// <param name="shape">The OLE-bearing shape. Must not be null.</param>
    /// <returns>
    ///     A <see cref="WordOleLocation" /> or <c>null</c> when the shape has no section
    ///     ancestor. <c>BodyParagraph</c> is <c>-1</c> when no body-paragraph ancestor
    ///     exists (e.g. header / footer).
    /// </returns>
    public static WordOleLocation? ResolveLocation(Document document, Shape shape)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(shape);

        var section = shape.GetAncestor(NodeType.Section) as Section;
        if (section == null) return null;

        var sectionIndex = document.Sections.IndexOf(section);
        var paragraph = shape.GetAncestor(NodeType.Paragraph) as Paragraph;
        var paragraphIndex = -1;
        if (paragraph != null)
        {
            var bodyParagraphs = section.Body?.Paragraphs;
            if (bodyParagraphs != null)
                paragraphIndex = bodyParagraphs.IndexOf(paragraph);
        }

        return new WordOleLocation(sectionIndex, paragraphIndex);
    }
}
