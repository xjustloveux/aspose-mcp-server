using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace AsposeMcpServer.Models;

public class McpResponse
{
    [JsonPropertyName("jsonrpc")]
    public string Jsonrpc { get; set; } = "2.0";
    
    [JsonPropertyName("id")]
    public JsonNode? Id { get; set; }
    
    [JsonPropertyName("result")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Result { get; set; }
    
    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public McpError? Error { get; set; }
}

public class McpError
{
    [JsonPropertyName("code")]
    public int Code { get; set; }
    
    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;
    
    [JsonPropertyName("data")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Data { get; set; }
}

