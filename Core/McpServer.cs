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
            PropertyNamingPolicy = null, // Don't use camelCase, use exact property names
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
                if (request == null) continue;

                var response = await HandleRequestAsync(request);
                
                // Only send response if not a notification
                if (response != null)
                {
                    var responseJson = JsonSerializer.Serialize(response, _jsonOptions);
                    await Console.Out.WriteLineAsync(responseJson);
                    await Console.Out.FlushAsync();
                }
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

    private async Task<McpResponse?> HandleRequestAsync(JsonObject request)
    {
        var method = request["method"]?.GetValue<string>();
        var id = request["id"];
        
        // Handle notifications (no id, no response needed)
        // MCP notifications include: initialized, and methods starting with "notifications/"
        if (id == null || (method == "initialized") || (method?.StartsWith("notifications/") == true))
        {
            Console.Error.WriteLine($"[DEBUG] Received notification: {method}");
            return null; // Notifications don't get a response
        }
        
        var response = new McpResponse
        {
            Jsonrpc = "2.0",
            Id = id
        };

        try
        {
            switch (method)
            {
                case "initialize":
                    // MCP protocolVersion uses YYYY-MM-DD format indicating the date of the last
                    // backward-incompatible change. This version should match the MCP specification.
                    // See: https://modelcontextprotocol.io/specification/2025-11-25
                    // 2025-11-25 version supports: tool annotations, pagination, tasks, cancellation
                    response.Result = new
                    {
                        protocolVersion = "2025-11-25",
                        serverInfo = new
                        {
                            name = "aspose-mcp-server",
                            version = "1.0.0"
                        },
                        capabilities = new
                        {
                            tools = new { }
                        }
                    };
                    break;

                case "tools/list":
                    // MCP 2025-11-25: tools/list response includes optional annotations
                    // Annotations provide metadata about tool behavior (e.g., readonly, destructive)
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
                            
                            // Get annotations: first check if tool implements IAnnotatedTool (manual override),
                            // otherwise infer from tool name pattern
                            var annotations = new Dictionary<string, object?>();
                            
                            if (tool is IAnnotatedTool annotatedTool)
                            {
                                // Manual annotations take precedence
                                if (annotatedTool.IsReadOnly.HasValue)
                                    annotations["readonly"] = annotatedTool.IsReadOnly.Value;
                                if (annotatedTool.IsDestructive.HasValue)
                                    annotations["destructive"] = annotatedTool.IsDestructive.Value;
                            }
                            else
                            {
                                // Auto-infer annotations from tool name patterns
                                var nameLower = toolName.ToLowerInvariant();
                                
                                // Read-only tools: get_*, extract_*, read, get_*, list_*
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
                            // Tool execution error - use centralized error handler
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

                default:
                    response.Result = null;
                    response.Error = McpErrorHandler.MethodNotFound(method ?? "null");
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[ERROR] Error handling method '{method}': {ex.Message}");
            
            // Only set error if not already set (e.g., from tools/call or default case)
            if (response.Error == null)
            {
                response.Result = null;
                response.Error = McpErrorHandler.HandleException(ex);
            }
        }

        return response;
    }
}

