using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Shared.Ole;

/// <summary>
///     Container-specific coordinates identifying where an OLE object lives inside its
///     host document. Emitted as a discriminated union with a <c>kind</c> tag so the
///     same cross-tool JSON shape (<see cref="OleObjectMetadata" />) can carry Word /
///     Excel / PowerPoint locations without schema inflation.
/// </summary>
[JsonPolymorphic(TypeDiscriminatorPropertyName = "kind")]
[JsonDerivedType(typeof(WordOleLocation), "word")]
[JsonDerivedType(typeof(ExcelOleLocation), "excel")]
[JsonDerivedType(typeof(PptOleLocation), "ppt")]
public abstract record OleShapeLocation;

/// <summary>
///     Location of an OLE object inside a Word document — the owning section's zero-based
///     index and the zero-based paragraph index within that section's body.
/// </summary>
/// <param name="Section">Zero-based section index in <c>Document.Sections</c>.</param>
/// <param name="BodyParagraph">
///     Zero-based paragraph index within the containing section's body. When the owning
///     shape is inside a header/footer or table cell, this is the nearest enclosing body
///     paragraph index, or <c>-1</c> when no body paragraph ancestor exists.
/// </param>
public sealed record WordOleLocation(
    [property: JsonPropertyName("section")]
    int Section,
    [property: JsonPropertyName("bodyParagraph")]
    int BodyParagraph) : OleShapeLocation;

/// <summary>
///     Location of an OLE object inside an Excel workbook — the owning worksheet name,
///     its zero-based index, and the upper-left anchor cell (row, column, zero-based).
/// </summary>
/// <param name="Sheet">Worksheet display name.</param>
/// <param name="SheetIndex">Zero-based worksheet index in <c>Workbook.Worksheets</c>.</param>
/// <param name="Row">Zero-based upper-left anchor row.</param>
/// <param name="Column">Zero-based upper-left anchor column.</param>
public sealed record ExcelOleLocation(
    [property: JsonPropertyName("sheet")] string Sheet,
    [property: JsonPropertyName("sheetIndex")]
    int SheetIndex,
    [property: JsonPropertyName("row")] int Row,
    [property: JsonPropertyName("column")] int Column) : OleShapeLocation;

/// <summary>
///     Location of an OLE object inside a PowerPoint presentation — the zero-based slide
///     index and the shape's zero-based position within that slide's shapes collection.
/// </summary>
/// <param name="SlideIndex">Zero-based slide index (<c>slide.SlideNumber - 1</c>).</param>
/// <param name="ShapeIndex">Zero-based shape index within <c>slide.Shapes</c>.</param>
public sealed record PptOleLocation(
    [property: JsonPropertyName("slideIndex")]
    int SlideIndex,
    [property: JsonPropertyName("shapeIndex")]
    int ShapeIndex) : OleShapeLocation;
