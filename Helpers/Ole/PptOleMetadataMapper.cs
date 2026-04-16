using Aspose.Slides;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Helpers.Ole;

/// <summary>
///     Projects an Aspose.Slides <see cref="IOleObjectFrame" /> into the cross-tool
///     <see cref="OleObjectMetadata" /> shape.
/// </summary>
public static class PptOleMetadataMapper
{
    /// <summary>
    ///     Maps a single PowerPoint OLE frame into the cross-tool metadata record.
    /// </summary>
    /// <param name="frame">The OLE frame. Must not be null.</param>
    /// <param name="slideIndex">Zero-based slide index.</param>
    /// <param name="shapeIndex">Zero-based shape index within the slide's shapes.</param>
    /// <param name="flatIndex">Zero-based flat index across the entire presentation.</param>
    /// <returns>A fully-populated <see cref="OleObjectMetadata" /> record.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="frame" /> is null.</exception>
    public static OleObjectMetadata Map(IOleObjectFrame frame, int slideIndex, int shapeIndex, int flatIndex)
    {
        ArgumentNullException.ThrowIfNull(frame);

        // ObjectProgId lives on the frame (api-verification.md item 2.b) — NOT on EmbeddedData.
        var progId = frame.ObjectProgId;
        long sizeBytes = 0;
        string? extension = null;

        // IOleEmbeddedDataInfo in Aspose.Slides 23.10.0 exposes ONLY EmbeddedFileData +
        // EmbeddedFileExtension (no filename); the frame carries EmbeddedFileName /
        // EmbeddedFileLabel separately.
        var rawName = !string.IsNullOrEmpty(frame.EmbeddedFileName)
            ? frame.EmbeddedFileName
            : frame.EmbeddedFileLabel;

        if (frame.EmbeddedData != null)
        {
            sizeBytes = frame.EmbeddedData.EmbeddedFileData?.Length ?? 0L;
            extension = OleSanitizerHelper.NormalizeExtension(frame.EmbeddedData.EmbeddedFileExtension);
        }

        if (string.IsNullOrEmpty(extension))
            extension = OleSanitizerHelper.ExtensionFromProgId(progId);

        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName(rawName, flatIndex, progId);

        string? linkTarget = null;
        if (frame.IsObjectLink)
        {
            var raw = frame.LinkPathLong;
            linkTarget = string.IsNullOrEmpty(raw) ? null : OleSanitizerHelper.SanitizeForLog(raw);
        }

        return new OleObjectMetadata
        {
            Index = flatIndex,
            RawFileName = rawName,
            SuggestedFileName = suggested,
            ProgId = progId,
            SizeBytes = sizeBytes,
            IsLinked = frame.IsObjectLink,
            Extension = extension,
            ShapeLocation = new PptOleLocation(slideIndex, shapeIndex),
            LinkTarget = linkTarget
        };
    }
}
