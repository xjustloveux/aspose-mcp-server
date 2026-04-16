using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Helpers.Ole;

/// <summary>
///     Projects an Aspose.Cells <see cref="OleObject" /> into the cross-tool
///     <see cref="OleObjectMetadata" /> shape.
/// </summary>
/// <remarks>
///     API note (resolving design §10 open item H, 2026-04-15): the Cells 23.10.0 type
///     <see cref="OleObject" /> does NOT expose a <c>FileName</c> property. Verified
///     against <c>Aspose.Cells.xml</c> — available string properties are
///     <c>SourceFullName</c> (marked <c>[Obsolete]</c>, replaced by
///     <c>ObjectSourceFullName</c>), <c>Label</c>, and <c>ImageSourceFullName</c>. We
///     use <c>Path.GetFileName(OleObject.ObjectSourceFullName)</c>, falling back to
///     <c>Label</c>; both are sanitized downstream.
/// </remarks>
public static class ExcelOleMetadataMapper
{
    /// <summary>
    ///     Maps a single Excel <see cref="OleObject" /> into the cross-tool metadata record.
    /// </summary>
    /// <param name="ole">The OLE object. Must not be null.</param>
    /// <param name="sheet">Owning worksheet. Must not be null.</param>
    /// <param name="sheetIndex">Zero-based worksheet index in <c>Workbook.Worksheets</c>.</param>
    /// <param name="flatIndex">Zero-based flat index across the entire workbook.</param>
    /// <returns>A fully-populated <see cref="OleObjectMetadata" /> record.</returns>
    /// <exception cref="ArgumentNullException">
    ///     Thrown when <paramref name="ole" /> or <paramref name="sheet" /> is null.
    /// </exception>
    public static OleObjectMetadata Map(OleObject ole, Worksheet sheet, int sheetIndex, int flatIndex)
    {
        ArgumentNullException.ThrowIfNull(ole);
        ArgumentNullException.ThrowIfNull(sheet);

        var progId = ole.ProgID;
        var rawName = ResolveRawFileName(ole);
        var (suggested, _) = OleSanitizerHelper.SanitizeOleFileName(rawName, flatIndex, progId);

        var extension = Path.GetExtension(rawName ?? string.Empty);
        extension = OleSanitizerHelper.NormalizeExtension(extension);
        if (string.IsNullOrEmpty(extension))
            extension = OleSanitizerHelper.ExtensionFromProgId(progId);

        var sizeBytes = ole.ObjectData?.Length ?? 0L;
        string? linkTarget = null;
        if (ole.IsLink)
        {
            var raw = ole.ObjectSourceFullName;
            linkTarget = string.IsNullOrEmpty(raw)
                ? null
                : OleSanitizerHelper.SanitizeForLog(raw);
        }

        var location = new ExcelOleLocation(sheet.Name, sheetIndex, ole.UpperLeftRow, ole.UpperLeftColumn);

        return new OleObjectMetadata
        {
            Index = flatIndex,
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
    ///     Resolves the best raw filename candidate: <c>Path.GetFileName(SourceFullName)</c>
    ///     (valid for linked OLE), <c>ObjectSourceFullName</c> (legacy alias), then
    ///     <c>Label</c> (user-visible display label). Embedded OLE frequently has empty
    ///     values for all three — in which case the sanitizer falls back to
    ///     <c>ole_&lt;index&gt;&lt;ext&gt;</c>.
    /// </summary>
    /// <param name="ole">The OLE object. Must not be null.</param>
    /// <returns>The first non-empty candidate, or <c>null</c>.</returns>
    private static string? ResolveRawFileName(OleObject ole)
    {
        // ObjectSourceFullName is the replacement for the obsolete SourceFullName
        // (Aspose.Cells 23.10.0 marks SourceFullName [Obsolete]).
        if (!string.IsNullOrEmpty(ole.ObjectSourceFullName))
            return Path.GetFileName(ole.ObjectSourceFullName);

        if (!string.IsNullOrEmpty(ole.Label))
            return ole.Label;

        return null;
    }
}
