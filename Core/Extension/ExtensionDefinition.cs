using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Represents a single extension definition from extensions.json.
/// </summary>
/// <remarks>
///     <para>
///         Extension metadata (name, version, description, author, websiteUrl) is obtained
///         dynamically through the initialization handshake protocol, not from the config file.
///     </para>
///     <para>
///         The ID is set from the dictionary key in extensions.json, not from a JSON property.
///     </para>
/// </remarks>
public class ExtensionDefinition
{
    /// <summary>
    ///     Unique identifier for the extension (set from dictionary key in extensions.json).
    /// </summary>
    [JsonIgnore]
    public string Id { get; set; } = string.Empty;

    /// <summary>
    ///     Runtime metadata obtained from handshake response. Null until handshake completes.
    /// </summary>
    [JsonIgnore]
    public ExtensionInitializeResponse? RuntimeMetadata { get; set; }

    /// <summary>
    ///     Gets the display name (from runtime metadata or fallback to ID).
    /// </summary>
    [JsonIgnore]
    public string DisplayName => RuntimeMetadata?.Name ?? Id;

    /// <summary>
    ///     Gets the version (from runtime metadata or "unknown").
    /// </summary>
    [JsonIgnore]
    public string DisplayVersion => RuntimeMetadata?.Version ?? "unknown";

    /// <summary>
    ///     Gets the title (from runtime metadata, for localized display).
    /// </summary>
    [JsonIgnore]
    public string? DisplayTitle => RuntimeMetadata?.Title;

    /// <summary>
    ///     Gets the description (from runtime metadata).
    /// </summary>
    [JsonIgnore]
    public string? DisplayDescription => RuntimeMetadata?.Description;

    /// <summary>
    ///     Gets the author (from runtime metadata).
    /// </summary>
    [JsonIgnore]
    public string? DisplayAuthor => RuntimeMetadata?.Author;

    /// <summary>
    ///     Gets the website URL (from runtime metadata).
    /// </summary>
    [JsonIgnore]
    public string? DisplayWebsiteUrl => RuntimeMetadata?.WebsiteUrl;

    /// <summary>
    ///     Command configuration for starting the extension process.
    /// </summary>
    [JsonPropertyName("command")]
    public ExtensionCommand Command { get; set; } = new();

    /// <summary>
    ///     Gets a value indicating whether this extension has a valid command configured.
    ///     Extensions without valid commands are used for definition/metadata only and cannot be started.
    /// </summary>
    [JsonIgnore]
    public bool HasValidCommand => !string.IsNullOrWhiteSpace(Command.Executable);

    /// <summary>
    ///     List of input formats the extension accepts (e.g., ["pdf", "html", "png"]).
    /// </summary>
    [JsonPropertyName("inputFormats")]
    public List<string> InputFormats { get; set; } = [];

    /// <summary>
    ///     List of document types the extension supports (e.g., ["word", "excel", "powerpoint", "pdf"]).
    /// </summary>
    [JsonPropertyName("supportedDocumentTypes")]
    public List<string> SupportedDocumentTypes { get; set; } = [];

    /// <summary>
    ///     List of transport modes the extension supports (e.g., ["mmap", "stdin", "file"]).
    /// </summary>
    [JsonPropertyName("transportModes")]
    public List<string> TransportModes { get; set; } = ["file"];

    /// <summary>
    ///     Preferred transport mode if multiple are supported.
    /// </summary>
    [JsonPropertyName("preferredTransportMode")]
    public string? PreferredTransportMode { get; set; }

    /// <summary>
    ///     Protocol version this extension implements.
    /// </summary>
    [JsonPropertyName("protocolVersion")]
    public string ProtocolVersion { get; set; } = "1.0";

    /// <summary>
    ///     Optional capabilities of the extension.
    /// </summary>
    [JsonPropertyName("capabilities")]
    public ExtensionCapabilities? Capabilities { get; set; }

    /// <summary>
    ///     Whether the extension is currently available (set at runtime based on validation).
    /// </summary>
    [JsonIgnore]
    public bool IsAvailable { get; set; } = true;

    /// <summary>
    ///     Error message if extension is not available (set at runtime).
    /// </summary>
    [JsonIgnore]
    public string? UnavailableReason { get; set; }

    /// <summary>
    ///     Gets the effective frame interval with constraints applied.
    /// </summary>
    /// <param name="config">The global extension configuration.</param>
    /// <returns>The effective frame interval in milliseconds, constrained to valid range.</returns>
    public int GetEffectiveFrameIntervalMs(ExtensionConfig config)
    {
        return config.FrameIntervalMs.Apply(Capabilities?.FrameIntervalMs);
    }

    /// <summary>
    ///     Gets the effective snapshot TTL with constraints applied.
    /// </summary>
    /// <param name="config">The global extension configuration.</param>
    /// <returns>The effective snapshot TTL in seconds, constrained to valid range.</returns>
    public int GetEffectiveSnapshotTtlSeconds(ExtensionConfig config)
    {
        return config.SnapshotTtlSeconds.Apply(Capabilities?.SnapshotTtlSeconds);
    }

    /// <summary>
    ///     Gets the effective max missed heartbeats with constraints applied.
    /// </summary>
    /// <param name="config">The global extension configuration.</param>
    /// <returns>The effective max missed heartbeats, constrained to valid range.</returns>
    public int GetEffectiveMaxMissedHeartbeats(ExtensionConfig config)
    {
        return config.MaxMissedHeartbeats.Apply(Capabilities?.MaxMissedHeartbeats);
    }

    /// <summary>
    ///     Gets the effective idle timeout with constraints and special value handling.
    ///     Special value 0 means "never unload" (permanent resident).
    /// </summary>
    /// <param name="config">The global extension configuration.</param>
    /// <returns>The effective idle timeout in minutes, or 0 for permanent resident if allowed.</returns>
    public int GetEffectiveIdleTimeoutMinutes(ExtensionConfig config)
    {
        return config.IdleTimeoutMinutes.Apply(Capabilities?.IdleTimeoutMinutes);
    }

    /// <summary>
    ///     Validates all extension capability settings against global constraints.
    ///     Returns warnings for any values that were constrained.
    /// </summary>
    /// <param name="config">The global extension configuration.</param>
    /// <returns>List of warning messages for constrained values. Empty if all values are within limits.</returns>
    public IReadOnlyList<string> ValidateCapabilityConstraints(ExtensionConfig config)
    {
        var warnings = new List<string>();

        var results = new[]
        {
            config.FrameIntervalMs.ApplyWithWarning(Capabilities?.FrameIntervalMs, "FrameIntervalMs"),
            config.SnapshotTtlSeconds.ApplyWithWarning(Capabilities?.SnapshotTtlSeconds, "SnapshotTtlSeconds"),
            config.MaxMissedHeartbeats.ApplyWithWarning(Capabilities?.MaxMissedHeartbeats, "MaxMissedHeartbeats"),
            config.IdleTimeoutMinutes.ApplyWithWarning(Capabilities?.IdleTimeoutMinutes, "IdleTimeoutMinutes")
        };

        foreach (var result in results)
            if (result.HasWarning)
                warnings.Add(result.Warning!);

        return warnings;
    }
}
