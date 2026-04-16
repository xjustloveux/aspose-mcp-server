using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Shared.Ole;

/// <summary>
///     Per-item entry emitted inside <see cref="OleExtractAllResult.Items" />, describing
///     one successfully extracted OLE object.
/// </summary>
public sealed record OleExtractAllItem
{
    /// <summary>Zero-based index of the OLE in the container snapshot.</summary>
    [JsonPropertyName("index")]
    [JsonPropertyOrder(1)]
    public int Index { get; init; }

    /// <summary>Absolute path to the written file.</summary>
    [JsonPropertyName("outputFilePath")]
    [JsonPropertyOrder(2)]
    public string OutputFilePath { get; init; } = string.Empty;

    /// <summary>Byte count written.</summary>
    [JsonPropertyName("bytesWritten")]
    [JsonPropertyOrder(3)]
    public long BytesWritten { get; init; }

    /// <summary>Whether the suggested filename differed from the raw filename.</summary>
    [JsonPropertyName("sanitizedFromRaw")]
    [JsonPropertyOrder(4)]
    public bool SanitizedFromRaw { get; init; }
}

/// <summary>
///     Result payload returned by the <c>extract_all</c> operation.
/// </summary>
public sealed record OleExtractAllResult
{
    /// <summary>Total OLE count in the container (including linked objects that are skipped).</summary>
    [JsonPropertyName("requested")]
    [JsonPropertyOrder(1)]
    public int Requested { get; init; }

    /// <summary>Number of items successfully written to disk.</summary>
    [JsonPropertyName("extracted")]
    [JsonPropertyOrder(2)]
    public int Extracted { get; init; }

    /// <summary>
    ///     Entries skipped with a machine-readable reason (<c>linked</c>, <c>empty-payload</c>,
    ///     <c>cumulative-size-cap-exceeded</c>, ...).
    /// </summary>
    [JsonPropertyName("skipped")]
    [JsonPropertyOrder(3)]
    public IReadOnlyList<OleSkippedEntry> Skipped { get; init; } = Array.Empty<OleSkippedEntry>();

    /// <summary>Ordered per-item entries for successfully extracted objects.</summary>
    [JsonPropertyName("items")]
    [JsonPropertyOrder(4)]
    public IReadOnlyList<OleExtractAllItem> Items { get; init; } = Array.Empty<OleExtractAllItem>();

    /// <summary>
    ///     <c>true</c> when the operation aborted early due to the cumulative-size cap
    ///     (F-8). Callers should inspect <see cref="Skipped" /> for the
    ///     <c>cumulative-size-cap-exceeded</c> entry marking where the cap kicked in.
    /// </summary>
    [JsonPropertyName("truncated")]
    [JsonPropertyOrder(5)]
    public bool Truncated { get; init; }

    /// <summary>See <see cref="OleListResult.PasswordIgnored" />.</summary>
    [JsonPropertyName("passwordIgnored")]
    [JsonPropertyOrder(6)]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public PasswordIgnoredNote? PasswordIgnored { get; init; }
}
