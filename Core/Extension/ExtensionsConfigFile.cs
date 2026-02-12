using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Root object for extensions.json configuration file.
/// </summary>
public class ExtensionsConfigFile
{
    /// <summary>
    ///     Schema version of the configuration file format.
    ///     Reserved for future version compatibility checks.
    /// </summary>
    /// <remarks>
    ///     <para>Current behavior:</para>
    ///     <list type="bullet">
    ///         <item>This field is optional and not currently validated</item>
    ///         <item>Future versions may use this for migration or compatibility</item>
    ///         <item>Recommended format: semantic version (e.g., "1.0", "1.1")</item>
    ///     </list>
    /// </remarks>
    [JsonPropertyName("schemaVersion")]
    public string? SchemaVersion { get; set; }

    /// <summary>
    ///     Dictionary of extension definitions, keyed by extension ID.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         The dictionary key serves as the unique extension ID.
    ///         This design prevents duplicate IDs at the JSON parsing level.
    ///     </para>
    ///     <para>
    ///         ID requirements:
    ///         <list type="bullet">
    ///             <item>Must not be empty or whitespace</item>
    ///             <item>Must not contain colons (:) as they are reserved</item>
    ///         </list>
    ///     </para>
    /// </remarks>
    [JsonPropertyName("extensions")]
    public Dictionary<string, ExtensionDefinition> Extensions { get; set; } = new();
}
