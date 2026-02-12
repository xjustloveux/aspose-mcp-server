using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Optional capabilities of an extension.
/// </summary>
/// <remarks>
///     <para>
///         Extension-specified values are constrained by global configuration limits.
///         Use <see cref="ExtensionDefinition.GetEffectiveFrameIntervalMs" /> and similar methods
///         to get the final constrained values.
///     </para>
/// </remarks>
public class ExtensionCapabilities
{
    /// <summary>
    ///     Whether the extension supports heartbeat messages.
    /// </summary>
    [JsonPropertyName("supportsHeartbeat")]
    public bool SupportsHeartbeat { get; set; } = true;

    /// <summary>
    ///     Requested frame interval in milliseconds between snapshot transmissions.
    ///     Will be constrained by global Floor and Ceiling limits.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         Lightweight extensions can handle frequent snapshots (e.g., 50ms),
    ///         while heavy processing extensions may need longer intervals (e.g., 500ms).
    ///     </para>
    /// </remarks>
    [JsonPropertyName("frameIntervalMs")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? FrameIntervalMs { get; set; }

    /// <summary>
    ///     Requested snapshot TTL in seconds for unacknowledged snapshots.
    ///     Will be constrained by global Floor and Ceiling limits.
    /// </summary>
    /// <remarks>
    ///     <para>Extensions that perform heavy processing may need longer TTL.</para>
    /// </remarks>
    [JsonPropertyName("snapshotTtlSeconds")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? SnapshotTtlSeconds { get; set; }

    /// <summary>
    ///     Requested maximum consecutive missed heartbeat responses before marking as unresponsive.
    ///     Will be constrained by global Floor and Ceiling limits.
    /// </summary>
    /// <remarks>
    ///     <para>Extensions with slow processing may need a higher tolerance (e.g., 5-10).</para>
    /// </remarks>
    [JsonPropertyName("maxMissedHeartbeats")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? MaxMissedHeartbeats { get; set; }

    /// <summary>
    ///     Requested idle timeout in minutes before unloading this extension.
    ///     Will be constrained by global Floor and Ceiling limits.
    ///     Special value 0 means "never unload" (permanent resident), allowed only if global config permits.
    /// </summary>
    /// <remarks>
    ///     <para>Use cases:</para>
    ///     <list type="bullet">
    ///         <item><c>null</c> - Use global default</item>
    ///         <item><c>0</c> - Never unload (if allowed by global config)</item>
    ///         <item><c>&gt;0</c> - Unload after N minutes of inactivity (constrained)</item>
    ///     </list>
    /// </remarks>
    [JsonPropertyName("idleTimeoutMinutes")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? IdleTimeoutMinutes { get; set; }
}
