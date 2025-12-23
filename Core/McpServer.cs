using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using AsposeMcpServer.Models;
using AsposeMcpServer.Tools;

namespace AsposeMcpServer.Core;

/// <summary>
///     MCP (Model Context Protocol) server implementation for Aspose document manipulation tools
///     Handles JSON-RPC requests and routes them to appropriate tools
/// </summary>
public class McpServer
{
    private readonly JsonSerializerOptions _jsonOptions;
    private readonly Dictionary<string, IAsposeTool> _tools;

    /// <summary>
    ///     Initializes a new instance of the MCP server with the specified configuration
    /// </summary>
    /// <param name="config">Server configuration including enabled tool categories</param>
    public McpServer(ServerConfig config)
    {
        _jsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy =
                JsonNamingPolicy.CamelCase, // Use camelCase for JSON property names (MCP protocol standard)
            WriteIndented = false,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            Encoder = JavaScriptEncoder
                .UnsafeRelaxedJsonEscaping // Support Unicode characters (Chinese, Japanese, etc.)
        };

        // Use automatic tool discovery instead of manual registration
        _tools = ToolRegistry.DiscoverTools(config);

        Console.Error.WriteLine($"[INFO] Registered {_tools.Count} tools using automatic discovery");
    }

    /// <summary>
    ///     Runs the MCP server, processing JSON-RPC requests from stdin and sending responses to stdout
    /// </summary>
    public async Task RunAsync()
    {
        await Console.Error.WriteLineAsync("[INFO] Aspose MCP Server started");
        await Console.Error.FlushAsync();

        while (true)
            try
            {
                var line = await Console.In.ReadLineAsync();

                if (string.IsNullOrEmpty(line)) break;

                var request = JsonSerializer.Deserialize<JsonObject>(line, _jsonOptions);
                if (request == null)
                {
                    await Console.Error.WriteLineAsync("[WARN] Failed to parse JSON request, skipping");
                    continue;
                }

                await HandleRequestAndSendResponseAsync(request);
            }
            catch (Exception ex)
            {
                await Console.Error.WriteLineAsync($"[ERROR] Error processing request: {ex.Message}");
                await Console.Error.WriteLineAsync($"[ERROR] Stack trace: {ex.StackTrace}");

                var errorResponse = new McpResponse
                {
                    Jsonrpc = "2.0",
                    Id = null,
                    Error = McpErrorHandler.ParseError(ex.Message)
                };
                var errorJson = JsonSerializer.Serialize(errorResponse, _jsonOptions);
                await Console.Out.WriteLineAsync(errorJson);
                await Console.Out.FlushAsync();
            }
    }

    private async Task HandleRequestAndSendResponseAsync(JsonObject request)
    {
        try
        {
            var method = request["method"]?.GetValue<string>();
            var id = request["id"];

            // Handle notifications (no id, no response needed)
            // MCP notifications include: initialized, and methods starting with "notifications/"
            if (id == null || method == "initialized" ||
                method?.StartsWith("notifications/") == true) return;

            var response = new McpResponse
            {
                Jsonrpc = "2.0",
                Id = id
            };

            if (method == "initialize")
            {
                var paramsObj = request["params"] as JsonObject;
                var clientProtocolVersion = paramsObj?["protocolVersion"]?.GetValue<string>();

                // MCP protocol: server should return the protocol version it will use
                // We support both 2025-06-18 and 2025-11-25, use client's version for compatibility
                var protocolVersion = clientProtocolVersion == "2025-06-18" ? "2025-06-18" : "2025-11-25";

                response.Result = new
                {
                    protocolVersion,
                    serverInfo = new
                    {
                        name = "aspose-mcp-server",
                        version = VersionHelper.GetVersion()
                    },
                    capabilities = new
                    {
                        tools = new { }
                    }
                };

                var responseJson = JsonSerializer.Serialize(response, _jsonOptions);
                await Console.Out.WriteLineAsync(responseJson);
                await Console.Out.FlushAsync();
                return;
            }

            try
            {
                await ProcessRequestAsync(request, response, method);
            }
            catch (Exception ex)
            {
                await Console.Error.WriteLineAsync($"[ERROR] Error handling method '{method}': {ex.Message}");

                if (response.Error == null)
                {
                    response.Result = null;
                    response.Error = McpErrorHandler.HandleException(ex);
                }
            }

            var responseJson2 = JsonSerializer.Serialize(response, _jsonOptions);
            await Console.Out.WriteLineAsync(responseJson2);
            await Console.Out.FlushAsync();
        }
        catch (Exception ex)
        {
            await Console.Error.WriteLineAsync(
                $"[ERROR] Fatal error in HandleRequestAndSendResponseAsync: {ex.Message}");
            await Console.Error.WriteLineAsync($"[ERROR] Stack trace: {ex.StackTrace}");
        }
    }

    private async Task ProcessRequestAsync(JsonObject request, McpResponse response, string? method)
    {
        switch (method)
        {
            case "tools/list":
                // MCP 2025-11-25: tools/list response includes optional annotations
                // Annotations provide metadata about tool behavior (readonly, destructive)
                response.Result = new
                {
                    tools = _tools.Select(kvp =>
                    {
                        var toolInstance = kvp.Value;
                        var toolName = kvp.Key;

                        var toolObj = new Dictionary<string, object?>
                        {
                            ["name"] = toolName,
                            ["description"] = toolInstance.Description,
                            ["inputSchema"] = toolInstance.InputSchema
                        };

                        // Get annotations: check IAnnotatedTool interface first, otherwise infer from name pattern
                        var annotations = new Dictionary<string, object?>();

                        if (toolInstance is IAnnotatedTool annotatedTool)
                        {
                            if (annotatedTool.IsReadOnly.HasValue)
                                annotations["readonly"] = annotatedTool.IsReadOnly.Value;
                            if (annotatedTool.IsDestructive.HasValue)
                                annotations["destructive"] = annotatedTool.IsDestructive.Value;
                        }
                        else
                        {
                            var nameLower = toolName.ToLowerInvariant();

                            // Read-only tools: get_*, extract_*, read, list_*, *_info, *_properties, etc.
                            if (nameLower.StartsWith("get_") ||
                                nameLower.StartsWith("extract_") ||
                                nameLower.StartsWith("read") ||
                                nameLower.StartsWith("list_") ||
                                nameLower.Contains("_get_") ||
                                nameLower.EndsWith("_info") ||
                                nameLower.EndsWith("_statistics") ||
                                nameLower.EndsWith("_properties") ||
                                nameLower.EndsWith("_details"))
                                annotations["readonly"] = true;

                            // Destructive tools: delete_*, remove_*, clear_*, split (may delete pages)
                            if (nameLower.StartsWith("delete_") ||
                                nameLower.StartsWith("remove_") ||
                                nameLower.StartsWith("clear_") ||
                                (nameLower == "split" && !nameLower.Contains("table")) ||
                                nameLower.Contains("_delete_") ||
                                nameLower.Contains("_remove_") ||
                                nameLower.Contains("_clear"))
                                annotations["destructive"] = true;
                        }

                        if (annotations.Count > 0)
                            toolObj["annotations"] = annotations;

                        return toolObj;
                    }).ToArray()
                };
                break;

            case "tools/call":
                var parameters = request["params"] as JsonObject;
                var toolName = parameters?["name"]?.GetValue<string>();
                var arguments = parameters?["arguments"] as JsonObject;

                if (string.IsNullOrEmpty(toolName))
                {
                    response.Result = null;
                    response.Error = new McpError
                    {
                        Code = -32602, // Invalid params
                        Message = "Tool name is required"
                    };
                    break;
                }

                if (_tools.TryGetValue(toolName, out var tool))
                {
                    try
                    {
                        var result = await tool.ExecuteAsync(arguments);
                        response.Result = new
                        {
                            content = new[]
                            {
                                new
                                {
                                    type = "text",
                                    text = result
                                }
                            }
                        };
                    }
                    catch (Exception ex)
                    {
                        response.Result = null;
                        response.Error = McpErrorHandler.HandleException(ex);
                        await Console.Error.WriteLineAsync($"[ERROR] Tool '{toolName}' execution failed: {ex.Message}");
                    }
                }
                else
                {
                    response.Result = null;
                    response.Error = McpErrorHandler.ToolNotFound(toolName);
                }

                break;

            case "ListOfferings":
            case "listOfferings":
                // Some MCP clients use ListOfferings to get server information
                response.Result = new
                {
                    serverInfo = new
                    {
                        name = "aspose-mcp-server",
                        version = VersionHelper.GetVersion()
                    },
                    protocolVersion = "2025-11-25",
                    capabilities = new
                    {
                        tools = new { }
                    },
                    tools = _tools.Keys.ToArray()
                };
                break;

            default:
                response.Result = null;
                response.Error = McpErrorHandler.MethodNotFound(method ?? "null");
                break;
        }
    }
}