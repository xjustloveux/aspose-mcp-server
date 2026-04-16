using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Results.Shared.Ole;
using Shape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Helpers.Ole;

/// <summary>
///     Cross-handler utility covering the small set of chores the 12 OLE handlers
///     perform identically — path validation, output-directory creation, index
///     resolution against the three Aspose collection shapes, and filename-override
///     sanitization. Extracted so cross-tool behavior stays lock-step (AC-18 / AC-19).
/// </summary>
public static class OleHandlerShared
{
    /// <summary>
    ///     Applies <see cref="SecurityHelper.ValidateFilePath" /> and, when a
    ///     <see cref="ServerConfig" /> is available, the allowlist check to an output
    ///     directory path. Raises <see cref="ArgumentException" /> on shape failure
    ///     and <see cref="UnauthorizedAccessException" /> on allowlist rejection —
    ///     both sanitized per F-5 / F-10 (no full paths leaked).
    /// </summary>
    /// <param name="outputDirectory">Caller-supplied directory path.</param>
    /// <param name="serverConfig">
    ///     Optional server config. When non-null, the directory must fall within
    ///     <see cref="ServerConfig.AllowedBasePaths" /> (if that list is configured).
    /// </param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the directory fails <see cref="SecurityHelper.ValidateFilePath" />.
    /// </exception>
    /// <exception cref="UnauthorizedAccessException">
    ///     Thrown when the directory falls outside the configured allowlist.
    /// </exception>
    public static void ValidateOutputDirectory(string outputDirectory, ServerConfig? serverConfig)
    {
        if (string.IsNullOrWhiteSpace(outputDirectory))
            throw new ArgumentException(OleErrorMessageBuilder.InvalidPath(outputDirectory), nameof(outputDirectory));

        try
        {
            SecurityHelper.ValidateFilePath(outputDirectory, nameof(outputDirectory), true);
        }
        catch (ArgumentException)
        {
            throw new ArgumentException(OleErrorMessageBuilder.InvalidPath(outputDirectory), nameof(outputDirectory));
        }

        if (serverConfig is { AllowedBasePaths.Count: > 0 })
            try
            {
                SecurityHelper.ResolveAndEnsureWithinAllowlist(outputDirectory, serverConfig.AllowedBasePaths,
                    nameof(outputDirectory));
            }
            catch (ArgumentException)
            {
                throw new UnauthorizedAccessException(
                    OleErrorMessageBuilder.OutputDirectoryNotWritable(outputDirectory));
            }
    }

    /// <summary>
    ///     Ensures the destination directory exists and is writable. Creates the directory
    ///     when absent (mirrors the precedent set by email / PDF attachment extractors).
    /// </summary>
    /// <param name="outputDirectory">Validated output directory path.</param>
    /// <exception cref="UnauthorizedAccessException">
    ///     Thrown when directory creation fails (permission, invalid name, I/O error).
    /// </exception>
    public static void EnsureDirectoryWritable(string outputDirectory)
    {
        try
        {
            Directory.CreateDirectory(outputDirectory);
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException
                                       or IOException
                                       or NotSupportedException
                                       or PathTooLongException
                                       or ArgumentException)
        {
            throw new UnauthorizedAccessException(
                OleErrorMessageBuilder.OutputDirectoryNotWritable(outputDirectory.TrimEnd(Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar)));
        }
    }

    /// <summary>
    ///     Picks the filename to write under — honoring the caller's <c>outputFileName</c>
    ///     override when supplied (still sanitized through
    ///     <see cref="OleSanitizerHelper.SanitizeOleFileName" />), else using the metadata's
    ///     <c>SuggestedFileName</c>.
    /// </summary>
    /// <param name="metadata">Sanitized cross-tool metadata for the target OLE object.</param>
    /// <param name="overrideName">
    ///     Optional caller-supplied filename (pre-sanitization). May be null / whitespace.
    /// </param>
    /// <returns>A disk-safe filename ready for <see cref="Path.Combine(string, string)" />.</returns>
    public static string ResolveExtractFileName(OleObjectMetadata metadata, string? overrideName)
    {
        if (!string.IsNullOrWhiteSpace(overrideName))
        {
            var (safe, _) = OleSanitizerHelper.SanitizeOleFileName(overrideName, metadata.Index, metadata.ProgId);
            return safe;
        }

        return metadata.SuggestedFileName;
    }

    /// <summary>
    ///     Resolves a zero-based flat index to the matching Word <see cref="Aspose.Words.Drawing.Shape" />.
    /// </summary>
    /// <param name="document">Source Word document. Must not be null.</param>
    /// <param name="flatIndex">Zero-based flat index across all OLE-bearing shapes.</param>
    /// <returns>A tuple of the matched shape and its flat index (echoed for callers).</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when no shape matches the index.</exception>
    public static (Shape shape, int index) LocateWordShape(Document document, int flatIndex)
    {
        ArgumentNullException.ThrowIfNull(document);

        var list = new List<Shape>();
        foreach (var shape in document.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            if (shape.OleFormat != null)
                list.Add(shape);

        if (flatIndex < 0 || flatIndex >= list.Count)
            throw new ArgumentOutOfRangeException(nameof(flatIndex), flatIndex,
                OleErrorMessageBuilder.IndexOutOfRange(flatIndex, list.Count));

        return (list[flatIndex], flatIndex);
    }

    /// <summary>
    ///     Resolves a zero-based flat index to a <c>(worksheet, localIndex)</c> tuple for
    ///     Excel. Index order is flat across sheets (sheet 0 all → sheet 1 all → …).
    /// </summary>
    /// <param name="workbook">Source workbook. Must not be null.</param>
    /// <param name="flatIndex">Zero-based flat index.</param>
    /// <returns>
    ///     Tuple <c>(ole, worksheet, sheetIndex, localIndex)</c> identifying the selected
    ///     OLE and its owning worksheet.
    /// </returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when no OLE object matches.</exception>
    public static (OleObject ole, Worksheet worksheet, int sheetIndex, int localIndex)
        LocateExcelOle(Workbook workbook, int flatIndex)
    {
        ArgumentNullException.ThrowIfNull(workbook);

        var cursor = 0;
        for (var si = 0; si < workbook.Worksheets.Count; si++)
        {
            var ws = workbook.Worksheets[si];
            var ols = ws.OleObjects;
            for (var li = 0; li < ols.Count; li++)
            {
                if (cursor == flatIndex)
                    return (ols[li], ws, si, li);
                cursor++;
            }
        }

        throw new ArgumentOutOfRangeException(nameof(flatIndex), flatIndex,
            OleErrorMessageBuilder.IndexOutOfRange(flatIndex, cursor));
    }

    /// <summary>
    ///     Resolves a zero-based flat index to a PowerPoint OLE frame plus its owning slide
    ///     / shape coordinates.
    /// </summary>
    /// <param name="presentation">Source presentation. Must not be null.</param>
    /// <param name="flatIndex">Zero-based flat index across all slides.</param>
    /// <returns>
    ///     Tuple <c>(frame, slide, slideIndex, shapeIndex)</c> identifying the selected
    ///     OLE frame and its location.
    /// </returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when no OLE frame matches.</exception>
    public static (IOleObjectFrame frame, ISlide slide, int slideIndex, int shapeIndex)
        LocatePptFrame(IPresentation presentation, int flatIndex)
    {
        ArgumentNullException.ThrowIfNull(presentation);

        var cursor = 0;
        for (var si = 0; si < presentation.Slides.Count; si++)
        {
            var slide = presentation.Slides[si];
            for (var shi = 0; shi < slide.Shapes.Count; shi++)
                if (slide.Shapes[shi] is IOleObjectFrame frame)
                {
                    if (cursor == flatIndex)
                        return (frame, slide, si, shi);
                    cursor++;
                }
        }

        throw new ArgumentOutOfRangeException(nameof(flatIndex), flatIndex,
            OleErrorMessageBuilder.IndexOutOfRange(flatIndex, cursor));
    }

    /// <summary>
    ///     Counts total OLE-bearing shapes in a Word document.
    /// </summary>
    /// <param name="document">Source document. Must not be null.</param>
    /// <returns>Non-negative count.</returns>
    public static int CountWordOle(Document document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var count = 0;
        foreach (var shape in document.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
            if (shape.OleFormat != null)
                count++;
        return count;
    }

    /// <summary>
    ///     Counts total OLE objects across all worksheets in an Excel workbook.
    /// </summary>
    /// <param name="workbook">Source workbook. Must not be null.</param>
    /// <returns>Non-negative count.</returns>
    public static int CountExcelOle(Workbook workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        var count = 0;
        foreach (var ws in workbook.Worksheets)
            count += ws.OleObjects.Count;
        return count;
    }

    /// <summary>
    ///     Counts total OLE frames across all slides in a presentation.
    /// </summary>
    /// <param name="presentation">Source presentation. Must not be null.</param>
    /// <returns>Non-negative count.</returns>
    public static int CountPptOle(IPresentation presentation)
    {
        ArgumentNullException.ThrowIfNull(presentation);
        var count = 0;
        foreach (var slide in presentation.Slides)
        foreach (var shape in slide.Shapes)
            if (shape is IOleObjectFrame)
                count++;
        return count;
    }
}
