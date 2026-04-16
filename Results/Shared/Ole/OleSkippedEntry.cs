using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Shared.Ole;

/// <summary>
///     Describes a single OLE object that <c>extract_all</c> skipped, with a machine-readable
///     reason. Recognized reasons include <c>linked</c>, <c>empty-payload</c>, and
///     <c>cumulative-size-cap-exceeded</c> (F-8).
/// </summary>
/// <param name="Index">Zero-based index of the skipped OLE in the container snapshot.</param>
/// <param name="Reason">Short machine-readable reason tag.</param>
public sealed record OleSkippedEntry(
    [property: JsonPropertyName("index")]
    [property: JsonPropertyOrder(1)]
    int Index,
    [property: JsonPropertyName("reason")]
    [property: JsonPropertyOrder(2)]
    string Reason);
