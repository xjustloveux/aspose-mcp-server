using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace AsposeMcpServer.Models;

/// <summary>
///     MCP JSON-RPC response model
/// </summary>
public class McpResponse
{
    /// <summary>
    ///     JSON-RPC protocol version (always "2.0")
    /// </summary>
    [JsonPropertyName("jsonrpc")]
    public string Jsonrpc { get; set; } = "2.0";

    /// <summary>
    ///     Request ID (null for notifications)
    /// </summary>
    [JsonPropertyName("id")]
    public JsonNode? Id { get; set; }

    /// <summary>
    ///     Response result (null if error occurred)
    /// </summary>
    [JsonPropertyName("result")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Result { get; set; }

    /// <summary>
    ///     Error object (null if request succeeded)
    /// </summary>
    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public McpError? Error { get; set; }
}

/// <summary>
///     MCP JSON-RPC error model
/// </summary>
public class McpError
{
    /// <summary>
    ///     Error code (negative number indicating error type)
    /// </summary>
    [JsonPropertyName("code")]
    public int Code { get; set; }

    /// <summary>
    ///     Error message describing what went wrong
    /// </summary>
    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    /// <summary>
    ///     Additional error data (optional)
    /// </summary>
    [JsonPropertyName("data")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Data { get; set; }
}