using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Command configuration for starting an extension process.
/// </summary>
/// <remarks>
///     <para>The command specifies how to launch the extension process. Supported command types:</para>
///     <list type="bullet">
///         <item><c>executable</c> - Direct executable (path in <see cref="Executable" />)</item>
///         <item><c>node</c> - Node.js script (uses <c>node</c> as interpreter)</item>
///         <item><c>python</c> - Python script (uses <c>python</c> or <c>python3</c>)</item>
///         <item><c>dotnet</c> - .NET assembly (uses <c>dotnet</c> runtime)</item>
///         <item><c>npx</c> - Node.js global package (uses <c>npx</c> to run)</item>
///         <item><c>pipx</c> - Python global package (uses <c>pipx run</c> to execute)</item>
///         <item><c>custom</c> - Custom command with full path to interpreter</item>
///     </list>
/// </remarks>
public class ExtensionCommand
{
    /// <summary>
    ///     Type of command: "executable", "node", "python", "dotnet", or "custom".
    /// </summary>
    [JsonPropertyName("type")]
    public string Type { get; set; } = "executable";

    /// <summary>
    ///     Path to the executable or script file.
    /// </summary>
    [JsonPropertyName("executable")]
    public string Executable { get; set; } = string.Empty;

    /// <summary>
    ///     Optional command line arguments.
    /// </summary>
    [JsonPropertyName("arguments")]
    public string? Arguments { get; set; }

    /// <summary>
    ///     Optional working directory for the process.
    /// </summary>
    [JsonPropertyName("workingDirectory")]
    public string? WorkingDirectory { get; set; }

    /// <summary>
    ///     Optional environment variables to set for the process.
    /// </summary>
    [JsonPropertyName("environment")]
    public Dictionary<string, string>? Environment { get; set; }
}
