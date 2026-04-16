using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Shared.Ole;

/// <summary>
///     Common projection of an OLE object across Word, Excel, and PowerPoint containers.
///     Shape is identical across all three tools so clients consuming
///     <c>word_ole_object</c>, <c>excel_ole_object</c>, and <c>ppt_ole_object</c> can
///     parse the result uniformly.
/// </summary>
public sealed record OleObjectMetadata
{
    /// <summary>
    ///     Zero-based flat index of the OLE object within the entire container (across
    ///     all sheets in Excel / all slides in PowerPoint). Stable only within a single
    ///     snapshot of the container; <c>remove</c> shifts subsequent indices down by one.
    /// </summary>
    [JsonPropertyName("index")]
    [JsonPropertyOrder(1)]
    public int Index { get; init; }

    /// <summary>
    ///     Raw filename as reported by the container (attacker-controlled, never written
    ///     to disk verbatim). Preserved for client-side display only. May be <c>null</c>
    ///     when the source carries no filename metadata.
    /// </summary>
    [JsonPropertyName("rawFileName")]
    [JsonPropertyOrder(2)]
    public string? RawFileName { get; init; }

    /// <summary>
    ///     Disk-safe filename derived from <see cref="RawFileName" /> via
    ///     <see cref="Helpers.Ole.OleSanitizerHelper.SanitizeOleFileName" />. This is the
    ///     name <c>extract</c> / <c>extract_all</c> will write under, absent a caller
    ///     override.
    /// </summary>
    [JsonPropertyName("suggestedFileName")]
    [JsonPropertyOrder(3)]
    public string SuggestedFileName { get; init; } = string.Empty;

    /// <summary>
    ///     ProgId reported by the container (e.g. <c>Excel.Sheet.12</c>). May be
    ///     <c>null</c> or empty for exotic payloads.
    /// </summary>
    [JsonPropertyName("progId")]
    [JsonPropertyOrder(4)]
    public string? ProgId { get; init; }

    /// <summary>
    ///     Uncompressed payload size in bytes. Zero when the OLE object is linked
    ///     (no embedded payload) or the size cannot be determined.
    /// </summary>
    [JsonPropertyName("sizeBytes")]
    [JsonPropertyOrder(5)]
    public long SizeBytes { get; init; }

    /// <summary>
    ///     <c>true</c> when the OLE object is a link to an external resource (no embedded
    ///     payload). <c>extract</c> / <c>extract_all</c> skip linked objects.
    /// </summary>
    [JsonPropertyName("isLinked")]
    [JsonPropertyOrder(6)]
    public bool IsLinked { get; init; }

    /// <summary>
    ///     Normalized extension (always dotted, e.g. <c>".xlsx"</c>). Defaults to
    ///     <c>".bin"</c> when neither the raw filename nor the ProgId yields a known
    ///     extension.
    /// </summary>
    [JsonPropertyName("extension")]
    [JsonPropertyOrder(7)]
    public string Extension { get; init; } = ".bin";

    /// <summary>
    ///     Container-specific location (<see cref="WordOleLocation" /> /
    ///     <see cref="ExcelOleLocation" /> / <see cref="PptOleLocation" />). May be
    ///     <c>null</c> when the location cannot be resolved (should not happen in
    ///     normal containers).
    /// </summary>
    [JsonPropertyName("shapeLocation")]
    [JsonPropertyOrder(8)]
    public OleShapeLocation? ShapeLocation { get; init; }

    /// <summary>
    ///     Sanitized link target (populated only when <see cref="IsLinked" /> is
    ///     <c>true</c>). Emitted as best-effort display metadata; never followed.
    /// </summary>
    [JsonPropertyName("linkTarget")]
    [JsonPropertyOrder(9)]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LinkTarget { get; init; }
}
