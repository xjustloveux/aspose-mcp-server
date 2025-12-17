using System.Text.RegularExpressions;
using AsposeMcpServer.Models;

namespace AsposeMcpServer.Core;

/// <summary>
///     Centralized error handler for MCP server errors
///     Provides consistent error handling and prevents information leakage
/// </summary>
public static class McpErrorHandler
{
    /// <summary>
    ///     Handles an exception and returns an appropriate MCP error response
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
                // Preserve ArgumentException details (parameter names, format examples) but remove file paths
                Message = SanitizeErrorMessage(argEx.Message, true)
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

            IOException => new McpError
            {
                Code = -32603, // Internal error
                Message = "I/O error occurred while processing the request"
            },

            InvalidOperationException opEx => new McpError
            {
                Code = -32603, // Internal error
                Message = SanitizeErrorMessage(opEx.Message)
            },

            NotSupportedException => new McpError
            {
                Code = -32603, // Internal error
                Message = "Operation not supported"
            },

            _ => new McpError
            {
                Code = -32603, // Internal error
                Message = includeStackTrace ? SanitizeErrorMessage(ex.Message) : "Internal server error occurred"
            }
        };
    }

    /// <summary>
    ///     Creates a method not found error
    /// </summary>
    /// <param name="method">Method name that was not found</param>
    /// <returns>MCP error object with method not found error code</returns>
    public static McpError MethodNotFound(string method)
    {
        return new McpError
        {
            Code = -32601, // Method not found
            Message = $"Unknown method: {method}"
        };
    }

    /// <summary>
    ///     Creates a tool not found error
    /// </summary>
    /// <param name="toolName">Tool name that was not found</param>
    /// <returns>MCP error object with method not found error code</returns>
    public static McpError ToolNotFound(string toolName)
    {
        return new McpError
        {
            Code = -32601, // Method not found
            Message = $"Unknown tool: {toolName}"
        };
    }

    /// <summary>
    ///     Creates a parse error
    /// </summary>
    /// <param name="message">Parse error message</param>
    /// <returns>MCP error object with parse error code</returns>
    public static McpError ParseError(string message)
    {
        return new McpError
        {
            Code = -32700, // Parse error
            Message = $"Parse error: {message}"
        };
    }

    /// <summary>
    ///     Creates an invalid params error
    /// </summary>
    /// <param name="message">Invalid parameters error message</param>
    /// <returns>MCP error object with invalid params error code</returns>
    public static McpError InvalidParams(string message)
    {
        return new McpError
        {
            Code = -32602, // Invalid params
            Message = SanitizeErrorMessage(message, true)
        };
    }

    /// <summary>
    ///     Sanitizes error messages to prevent information leakage
    ///     Removes file paths, stack traces, and other sensitive information
    /// </summary>
    /// <param name="message">Error message to sanitize</param>
    /// <param name="preserveDetails">If true, preserves detailed error messages (for ArgumentException)</param>
    /// <returns>Sanitized error message</returns>
    private static string SanitizeErrorMessage(string message, bool preserveDetails = false)
    {
        if (string.IsNullOrWhiteSpace(message)) return "An error occurred";

        var sanitized = message;

        if (!preserveDetails)
        {
            // Remove absolute file paths
            sanitized = Regex.Replace(sanitized, @"[A-Za-z]:\\[^\s]+", "[path removed]");
            sanitized = Regex.Replace(sanitized, @"/[^\s]+", "[path removed]");

            // Remove stack trace indicators
            sanitized = Regex.Replace(sanitized, @"at\s+[^\r\n]+", "");
            sanitized = Regex.Replace(sanitized, @"in\s+[^\r\n]+", "");
            sanitized = Regex.Replace(sanitized, @"line\s+\d+", "");

            // Remove exception type names that might leak implementation details
            sanitized = Regex.Replace(sanitized, @"\w+\.\w+Exception", "Error");
        }
        else
        {
            // For ArgumentException, only remove absolute file paths, preserve detailed messages
            sanitized = Regex.Replace(sanitized, @"[A-Za-z]:\\[^\s]+", "[path removed]");
            sanitized = Regex.Replace(sanitized, @"\/[^\s]+", "[path removed]");
        }

        // Limit message length (allow longer messages for detailed errors)
        var maxLength = preserveDetails ? 2000 : 500;
        if (sanitized.Length > maxLength) sanitized = sanitized.Substring(0, maxLength - 3) + "...";

        return sanitized.Trim();
    }
}