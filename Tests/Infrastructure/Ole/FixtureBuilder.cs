using Aspose.Cells;
using Aspose.Slides;
using Aspose.Slides.DOM.Ole;
using Aspose.Words;
using SkiaSharp;
using WordSaveFormat = Aspose.Words.SaveFormat;
using CellsSaveFormat = Aspose.Cells.SaveFormat;
using Shape = Aspose.Words.Drawing.Shape;
using SlidesSaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Infrastructure.Ole;

/// <summary>
///     Generates the 12-fixture matrix for the OLE-extract-unified test suite in-process
///     using Aspose APIs at assembly initialization time (per design §5.1). Fixtures live
///     under a unique subdirectory of <see cref="Path.GetTempPath" /> to avoid polluting
///     the source tree and are torn down at process exit. See
///     <see cref="OleFixtureCollection" /> for the xUnit collection-fixture wiring (built by
///     the test-engineer in a follow-up stage).
/// </summary>
public sealed class FixtureBuilder : IDisposable
{
    private readonly Dictionary<FixtureKind, string> _paths = new();

    /// <summary>
    ///     Initializes a new fixture directory and generates all 12 fixtures. Throws on
    ///     any Aspose / IO failure.
    /// </summary>
    public FixtureBuilder()
    {
        Root = Path.Combine(Path.GetTempPath(), "AsposeMcpOleFixtures_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Root);
        Directory.CreateDirectory(Path.Combine(Root, "word"));
        Directory.CreateDirectory(Path.Combine(Root, "excel"));
        Directory.CreateDirectory(Path.Combine(Root, "powerpoint"));

        GenerateWord();
        GenerateExcel();
        GeneratePpt();
    }

    /// <summary>Absolute path to the fixture root.</summary>
    public string Root { get; }

    /// <summary>Full path to each generated fixture, keyed by <see cref="FixtureKind" />.</summary>
    public IReadOnlyDictionary<FixtureKind, string> Paths => _paths;

    /// <summary>Deletes the fixture directory tree best-effort.</summary>
    public void Dispose()
    {
        try
        {
            if (Directory.Exists(Root)) Directory.Delete(Root, true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            /* best-effort cleanup */
        }
    }

    /// <summary>
    ///     Builds a small in-memory xlsx payload used as the embedded OLE data across
    ///     all three container generators.
    /// </summary>
    /// <returns>Byte array containing a valid .xlsx stream.</returns>
    private static byte[] BuildXlsxPayload()
    {
        using var wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("OLE fixture payload");
        using var ms = new MemoryStream();
        wb.Save(ms, CellsSaveFormat.Xlsx);
        return ms.ToArray();
    }

    /// <summary>
    ///     Builds a tiny PNG used as the PowerPoint OLE substitute picture. Returning raw
    ///     bytes keeps the generator independent of the Aspose.Slides image-load path.
    /// </summary>
    /// <returns>PNG byte array.</returns>
    private static byte[] BuildSubstitutePng()
    {
        using var bmp = new SKBitmap(8, 8);
        using var img = SKImage.FromBitmap(bmp);
        using var data = img.Encode(SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }

    /// <summary>Generates the four Word fixtures (docx + doc, embedded + linked + attacker + legacy).</summary>
    private void GenerateWord()
    {
        var xlsx = BuildXlsxPayload();

        _paths[FixtureKind.WordEmbeddedDocx] = BuildWord(xlsx, Path.Combine(Root, "word", "ole_embedded.docx"),
            WordSaveFormat.Docx, "payload.xlsx", false);
        _paths[FixtureKind.WordLinkedDocx] = BuildWord(xlsx, Path.Combine(Root, "word", "ole_linked.docx"),
            WordSaveFormat.Docx, "linked.xlsx", true);
        _paths[FixtureKind.WordAttackerDocx] = BuildWord(xlsx, Path.Combine(Root, "word", "ole_attacker_filename.docx"),
            WordSaveFormat.Docx, "..\\..\\etc\\passwd", false);
        _paths[FixtureKind.WordEmbeddedDoc] = BuildWord(xlsx, Path.Combine(Root, "word", "ole_embedded_legacy.doc"),
            WordSaveFormat.Doc, "legacy.xlsx", false);
    }

    /// <summary>Builds a Word document with a single OLE package attached.</summary>
    /// <param name="payload">Raw payload bytes to embed.</param>
    /// <param name="path">Full output path.</param>
    /// <param name="format">Save format (<c>Docx</c> or <c>Doc</c>).</param>
    /// <param name="rawName">Filename written into the OLE package metadata.</param>
    /// <param name="linked">When <c>true</c>, tweak post-save metadata so <c>IsLink</c> is true.</param>
    /// <returns>The written path (for convenience).</returns>
    private static string BuildWord(byte[] payload, string path, WordSaveFormat format, string rawName, bool linked)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        if (linked)
        {
            // Aspose.Words treats InsertOleObject(Stream, progId, asIcon, presentation)'s
            // third arg as asIcon, NOT as linked — to author a truly linked OLE we must
            // go through the filename overload, which requires a real sibling payload
            // file on disk. Write the xlsx payload to a sidecar path so the linked
            // reference resolves at Save() time.
            var sidecar = path + ".linked_payload.xlsx";
            File.WriteAllBytes(sidecar, payload);
            using var imgMsLinked = new MemoryStream(BuildSubstitutePng());
            builder.InsertOleObject(sidecar, "Excel.Sheet.12", true, false, imgMsLinked);

            var shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            if (shape.OleFormat != null)
                shape.OleFormat.SourceFullName = rawName;

            doc.Save(path, format);
            try
            {
                File.Delete(sidecar);
            }
            catch (IOException)
            {
                /* best-effort cleanup */
            }

            return path;
        }

        // Aspose.Words requires a non-null image (rendering substitute) for embedded OLE
        // — the 4th parameter name in the XML doc is "imageStream", not "presentation"
        // as the type signature suggests. Supply the 1x1 PNG built by BuildSubstitutePng.
        using var ms = new MemoryStream(payload);
        using var imgMs = new MemoryStream(BuildSubstitutePng());
        builder.InsertOleObject(ms, "Excel.Sheet.12", false, imgMs);

        var embedded = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (embedded.OleFormat?.OlePackage != null)
            embedded.OleFormat.OlePackage.FileName = rawName;

        doc.Save(path, format);
        return path;
    }

    /// <summary>Generates the four Excel fixtures.</summary>
    private void GenerateExcel()
    {
        var xlsx = BuildXlsxPayload();
        var png = BuildSubstitutePng();

        _paths[FixtureKind.ExcelEmbeddedXlsx] = BuildExcel(xlsx, png,
            Path.Combine(Root, "excel", "ole_embedded.xlsx"), CellsSaveFormat.Xlsx, "payload.xlsx", false);
        _paths[FixtureKind.ExcelLinkedXlsx] = BuildExcel(xlsx, png,
            Path.Combine(Root, "excel", "ole_linked.xlsx"), CellsSaveFormat.Xlsx, "linked.xlsx", true);
        _paths[FixtureKind.ExcelAttackerXlsx] = BuildExcel(xlsx, png,
            Path.Combine(Root, "excel", "ole_attacker_filename.xlsx"), CellsSaveFormat.Xlsx,
            "..\\..\\etc\\passwd", false);
        _paths[FixtureKind.ExcelEmbeddedXls] = BuildExcel(xlsx, png,
            Path.Combine(Root, "excel", "ole_embedded_legacy.xls"), CellsSaveFormat.Excel97To2003,
            "legacy.xlsx", false);
    }

    /// <summary>Builds an Excel workbook with a single OLE object on sheet 0.</summary>
    /// <param name="payload">Payload bytes.</param>
    /// <param name="substitutePng">Display-picture bytes.</param>
    /// <param name="path">Full output path.</param>
    /// <param name="format">Save format (<c>Xlsx</c> or <c>Excel97To2003</c>).</param>
    /// <param name="label">Label stored on the OLE object (raw filename surrogate).</param>
    /// <param name="linked">When <c>true</c>, set <c>ObjectSourceFullName</c> to mark the OLE linked.</param>
    /// <returns>The written path.</returns>
    private static string BuildExcel(byte[] payload, byte[] substitutePng, string path, CellsSaveFormat format,
        string label, bool linked)
    {
        using var wb = new Workbook();
        var sheet = wb.Worksheets[0];
        var idx = sheet.OleObjects.Add(1, 1, 100, 100, substitutePng);
        var ole = sheet.OleObjects[idx];
        ole.ProgID = "Excel.Sheet.12";
        ole.ObjectData = payload;
        ole.Label = label;
        if (linked)
            ole.ObjectSourceFullName = label;

        wb.Save(path, format);
        return path;
    }

    /// <summary>Generates the four PowerPoint fixtures.</summary>
    private void GeneratePpt()
    {
        var xlsx = BuildXlsxPayload();

        _paths[FixtureKind.PptEmbeddedPptx] = BuildPpt(xlsx,
            Path.Combine(Root, "powerpoint", "ole_embedded.pptx"), SlidesSaveFormat.Pptx,
            "payload.xlsx", false);
        _paths[FixtureKind.PptLinkedPptx] = BuildPpt(xlsx,
            Path.Combine(Root, "powerpoint", "ole_linked.pptx"), SlidesSaveFormat.Pptx,
            "linked.xlsx", true);
        _paths[FixtureKind.PptAttackerPptx] = BuildPpt(xlsx,
            Path.Combine(Root, "powerpoint", "ole_attacker_filename.pptx"), SlidesSaveFormat.Pptx,
            "..\\..\\etc\\passwd", false);
        _paths[FixtureKind.PptEmbeddedPpt] = BuildPpt(xlsx,
            Path.Combine(Root, "powerpoint", "ole_embedded_legacy.ppt"), SlidesSaveFormat.Ppt,
            "legacy.xlsx", false);
    }

    /// <summary>Builds a PowerPoint presentation with a single OLE frame on slide 1.</summary>
    /// <param name="payload">Payload bytes.</param>
    /// <param name="path">Full output path.</param>
    /// <param name="format">Save format (<c>Pptx</c> or <c>Ppt</c>).</param>
    /// <param name="embeddedName">Filename stored on the OLE frame.</param>
    /// <param name="linked">When <c>true</c>, set <c>LinkPathLong</c> on the frame.</param>
    /// <returns>The written path.</returns>
    private static string BuildPpt(byte[] payload, string path, SlidesSaveFormat format,
        string embeddedName, bool linked)
    {
        using var pres = new Presentation();
        var slide = pres.Slides[0];

        if (linked)
        {
            // Linked-OLE construction: the string-overload of AddOleObjectFrame takes
            // a progId + link path and flips IsObjectLink=true on the resulting frame.
            // LinkPathLong then carries the attacker-controlled filename through to the
            // sanitizer (PPT's only writable raw-filename side channel in 23.10.0).
            var linkTarget = Path.Combine(Path.GetDirectoryName(path)!, "linked_target.xlsx");
            File.WriteAllBytes(linkTarget, payload);
            slide.Shapes.AddOleObjectFrame(10, 10, 200, 150, "Excel.Sheet.12", linkTarget);
            var linkedFrame = (IOleObjectFrame)slide.Shapes[0];
            linkedFrame.LinkPathLong = embeddedName;
            pres.Save(path, format);
            try
            {
                File.Delete(linkTarget);
            }
            catch (IOException)
            {
                /* best-effort cleanup */
            }

            return path;
        }

        // Embedded case: AddOleObjectFrame(Single,Single,Single,Single,IOleEmbeddedDataInfo).
        // Aspose.Slides 23.10.0: EmbeddedFileName AND EmbeddedFileLabel are both read-only
        // on IOleObjectFrame (implementer note), so embedded attacker-filename fixtures
        // exercise the sanitizer via the empty-raw-name fallback path (ole_N.xlsx).
        var dataInfo = new OleEmbeddedDataInfo(payload, "xlsx");
        slide.Shapes.AddOleObjectFrame(10, 10, 200, 150, dataInfo);
        _ = embeddedName;
        pres.Save(path, format);
        return path;
    }
}
