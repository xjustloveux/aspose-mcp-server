using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using AsposeMcpServer.Models;
using AsposeMcpServer.Tools;

namespace AsposeMcpServer.Core;

public class McpServer
{
    private readonly Dictionary<string, IAsposeTool> _tools;
    private readonly JsonSerializerOptions _jsonOptions;
    private readonly ServerConfig _config;

    public McpServer(ServerConfig config)
    {
        _config = config;
        _jsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase, // Use camelCase for JSON property names (MCP protocol standard)
            WriteIndented = false,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping // Support Unicode characters (Chinese, Japanese, etc.)
        };

        // Use automatic tool discovery instead of manual registration
        // This reduces code from 600+ lines to a single line and makes maintenance easier
        _tools = ToolRegistry.DiscoverTools(_config);
        
        Console.Error.WriteLine($"[INFO] Registered {_tools.Count} tools using automatic discovery");
    }

    public async Task RunAsync()
    {
            Console.Error.WriteLine("[INFO] Aspose MCP Server started");
        await Console.Error.FlushAsync();
        
        while (true)
        {
            try
            {
                var line = await Console.In.ReadLineAsync();
                
                if (string.IsNullOrEmpty(line))
                {
                    break;
                }

                var request = JsonSerializer.Deserialize<JsonObject>(line, _jsonOptions);
                if (request == null)
                {
                    Console.Error.WriteLine("[WARN] Failed to parse JSON request, skipping");
                    continue;
                }

            await HandleRequestAndSendResponseAsync(request);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[ERROR] Error processing request: {ex.Message}");
                Console.Error.WriteLine($"[ERROR] Stack trace: {ex.StackTrace}");
                
                // For parse errors, we might not have a valid request ID
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
    }

    private async Task HandleRequestAndSendResponseAsync(JsonObject request)
    {
        try
        {
            var method = request["method"]?.GetValue<string>();
            var id = request["id"];
            
            // Handle notifications (no id, no response needed)
            // MCP notifications include: initialized, and methods starting with "notifications/"
            if (id == null || (method == "initialized") || (method?.StartsWith("notifications/") == true))
            {
                return; // Notifications don't get a response
            }
            
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
                    protocolVersion = protocolVersion,
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
                Console.Error.WriteLine($"[ERROR] Error handling method '{method}': {ex.Message}");
                
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
            Console.Error.WriteLine($"[ERROR] Fatal error in HandleRequestAndSendResponseAsync: {ex.Message}");
            Console.Error.WriteLine($"[ERROR] Stack trace: {ex.StackTrace}");
        }
    }

    private async Task ProcessRequestAsync(JsonObject request, McpResponse response, string? method)
    {
        switch (method)
            {
                case "initialize":
                    // This case is handled earlier in HandleRequestAndSendResponseAsync
                    // MCP protocolVersion uses YYYY-MM-DD format (date of last backward-incompatible change)
                    // See: https://modelcontextprotocol.io/specification/2025-11-25
                    // 2025-11-25 version supports: tool annotations, pagination, tasks, cancellation
                    response.Result = new
                    {
                        protocolVersion = "2025-11-25",
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
                    break;

                case "tools/list":
                    // MCP 2025-11-25: tools/list response includes optional annotations
                    // Annotations provide metadata about tool behavior (readonly, destructive)
                    response.Result = new
                    {
                        tools = _tools.Select(kvp =>
                        {
                            var tool = kvp.Value;
                            var toolName = kvp.Key;
                            
                            var toolObj = new Dictionary<string, object?>
                            {
                                ["name"] = toolName,
                                ["description"] = tool.Description,
                                ["inputSchema"] = tool.InputSchema
                            };
                            
                            // Get annotations: check IAnnotatedTool interface first, otherwise infer from name pattern
                            var annotations = new Dictionary<string, object?>();
                            
                            if (tool is IAnnotatedTool annotatedTool)
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
                                {
                                    annotations["readonly"] = true;
                                }
                                
                                // Destructive tools: delete_*, remove_*, clear_*, split (may delete pages)
                                if (nameLower.StartsWith("delete_") || 
                                    nameLower.StartsWith("remove_") || 
                                    nameLower.StartsWith("clear_") ||
                                    (nameLower == "split" && !nameLower.Contains("table")) ||
                                    nameLower.Contains("_delete_") ||
                                    nameLower.Contains("_remove_") ||
                                    nameLower.Contains("_clear"))
                                {
                                    annotations["destructive"] = true;
                                }
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
                            Console.Error.WriteLine($"[ERROR] Tool '{toolName}' execution failed: {ex.Message}");
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

