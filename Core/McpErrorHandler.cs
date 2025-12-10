using AsposeMcpServer.Models;

namespace AsposeMcpServer.Core;

/// <summary>
/// Centralized error handler for MCP server errors
/// Provides consistent error handling and prevents information leakage
/// </summary>
public static class McpErrorHandler
{
    /// <summary>
    /// Handles an exception and returns an appropriate MCP error response
    /// </summary>
    /// <param name="ex">The exception to handle</param>
    /// <param name="includeStackTrace">Whether to include stack trace in error message (for debugging)</param>
    /// <returns>An MCP error object</returns>
    public static McpError HandleException(Exception ex, bool includeStackTrace = false)
    {
        return ex switch
        {
            ArgumentNullException => new McpError
            {
                Code = -32602, // Invalid params
                Message = "Required parameter is missing or null"
            },
            
            ArgumentException argEx => new McpError
            {
                Code = -32602, // Invalid params
                Message = argEx.Message
            },
            
            FileNotFoundException => new McpError
            {
                Code = -32602, // Invalid params
                Message = "File not found"
            },
            
            DirectoryNotFoundException => new McpError
            {
                Code = -32602, // Invalid params
                Message = "Directory not found"
            },
            
            UnauthorizedAccessException => new McpError
            {
                Code = -32603, // Internal error
                Message = "Access denied to file or directory"
            },
            
            IOException ioEx => new McpError
            {
                Code = -32603, // Internal error
                Message = $"I/O error: {ioEx.Message}"
            },
            
            InvalidOperationException opEx => new McpError
            {
                Code = -32603, // Internal error
                Message = opEx.Message
            },
            
            NotSupportedException => new McpError
            {
                Code = -32603, // Internal error
                Message = "Operation not supported"
            },
            
            _ => new McpError
            {
                Code = -32603, // Internal error
                Message = includeStackTrace ? ex.ToString() : "Internal server error occurred"
            }
        };
    }
    
    /// <summary>
    /// Creates a method not found error
    /// </summary>
    public static McpError MethodNotFound(string method)
    {
        return new McpError
        {
            Code = -32601, // Method not found
            Message = $"Unknown method: {method}"
        };
    }
    
    /// <summary>
    /// Creates a tool not found error
    /// </summary>
    public static McpError ToolNotFound(string toolName)
    {
        return new McpError
        {
            Code = -32601, // Method not found
            Message = $"Unknown tool: {toolName}"
        };
    }
    
    /// <summary>
    /// Creates a parse error
    /// </summary>
    public static McpError ParseError(string message)
    {
        return new McpError
        {
            Code = -32700, // Parse error
            Message = $"Parse error: {message}"
        };
    }
    
    /// <summary>
    /// Creates an invalid params error
    /// </summary>
    public static McpError InvalidParams(string message)
    {
        return new McpError
        {
            Code = -32602, // Invalid params
            Message = message
        };
    }
}

