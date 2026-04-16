using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Shared.Ole;

/// <summary>
///     Result payload returned by the <c>remove</c> operation.
/// </summary>
public sealed record OleRemoveResult
{
    /// <summary>Zero-based index of the OLE object that was removed.</summary>
    [JsonPropertyName("index")]
    [JsonPropertyOrder(1)]
    public int Index { get; init; }

    /// <summary>Always <c>true</c> when the handler returns successfully.</summary>
    [JsonPropertyName("removed")]
    [JsonPropertyOrder(2)]
    public bool Removed { get; init; }

    /// <summary>
    ///     Absolute path of the re-saved container in file-mode. Absent in session-mode
    ///     where the change lives in the in-memory session.
    /// </summary>
    [JsonPropertyName("savedTo")]
    [JsonPropertyOrder(3)]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SavedTo { get; init; }

    /// <summary>See <see cref="OleListResult.PasswordIgnored" />.</summary>
    [JsonPropertyName("passwordIgnored")]
    [JsonPropertyOrder(4)]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public PasswordIgnoredNote? PasswordIgnored { get; init; }
}
