using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Shared.Ole;

/// <summary>
///     Result payload returned by the <c>list</c> operation. Shape is identical across
///     the three OLE tools (AC-11 / AC-18).
/// </summary>
public sealed record OleListResult
{
    /// <summary>
    ///     Total OLE object count in the container (including linked and embedded).
    /// </summary>
    [JsonPropertyName("count")]
    [JsonPropertyOrder(1)]
    public int Count { get; init; }

    /// <summary>
    ///     Reserved for forward compatibility. Always <c>false</c> in v1 (no hard cap;
    ///     see spec §scope_in point 8).
    /// </summary>
    [JsonPropertyName("truncated")]
    [JsonPropertyOrder(2)]
    public bool Truncated { get; init; }

    /// <summary>
    ///     Ordered list of metadata entries, one per OLE object. Preserves flat container
    ///     order (across sheets in Excel, across slides in PowerPoint).
    /// </summary>
    [JsonPropertyName("items")]
    [JsonPropertyOrder(3)]
    public IReadOnlyList<OleObjectMetadata> Items { get; init; } = Array.Empty<OleObjectMetadata>();

    /// <summary>
    ///     Session-mode advisory note emitted when the caller supplied a <c>password</c>
    ///     parameter that was ignored because the session was already unlocked. Absent in
    ///     file-mode responses.
    /// </summary>
    [JsonPropertyName("passwordIgnored")]
    [JsonPropertyOrder(4)]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public PasswordIgnoredNote? PasswordIgnored { get; init; }
}

/// <summary>
///     Locked-shape advisory note for the session-mode password-ignored case (F-5). The
///     name <c>passwordIgnored: true</c> plus a fixed reason string is the only content —
///     no <c>value</c>, no <c>attempted</c>, no echo of the password.
/// </summary>
public sealed record PasswordIgnoredNote
{
    /// <summary>
    ///     Always <c>true</c> on the wire. Present so JSON consumers can test the flag
    ///     directly rather than null-check the outer object.
    /// </summary>
    [JsonPropertyName("passwordIgnored")]
    [JsonPropertyOrder(1)]
    public bool PasswordIgnored { get; init; } = true;

    /// <summary>
    ///     Fixed machine-readable reason string (<c>"session-already-unlocked"</c>).
    /// </summary>
    [JsonPropertyName("reason")]
    [JsonPropertyOrder(2)]
    public string Reason { get; init; } = "session-already-unlocked";
}
