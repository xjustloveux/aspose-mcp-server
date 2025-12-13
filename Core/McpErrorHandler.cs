using System.Text.RegularExpressions;
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
                // CRITICAL FIX: Don't sanitize ArgumentException messages - they contain helpful user-facing error details
                // Sanitization was removing important information like parameter names and format examples
                // Only sanitize to remove file paths, but keep the detailed error message
                Message = SanitizeErrorMessage(argEx.Message, preserveDetails: true)
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
                // Don't expose internal I/O error details to prevent information leakage
                Message = "I/O error occurred while processing the request"
            },
            
            InvalidOperationException opEx => new McpError
            {
                Code = -32603, // Internal error
                // Only expose user-friendly messages, sanitize internal details
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
                // Never expose stack traces or internal error details in production
                Message = includeStackTrace ? SanitizeErrorMessage(ex.Message) : "Internal server error occurred"
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
            Message = SanitizeErrorMessage(message, preserveDetails: true)
        };
    }

    /// <summary>
    /// Sanitizes error messages to prevent information leakage
    /// Removes file paths, stack traces, and other sensitive information
    /// </summary>
    /// <param name="preserveDetails">If true, preserves detailed error messages (for ArgumentException)</param>
    private static string SanitizeErrorMessage(string message, bool preserveDetails = false)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return "An error occurred";
        }

        var sanitized = message;
        
        if (!preserveDetails)
        {
            // Remove file paths (basic pattern matching) - but preserve relative paths and filenames
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
            // For ArgumentException, only remove absolute file paths, preserve everything else
            // This keeps detailed error messages with parameter names, format examples, etc.
            sanitized = Regex.Replace(sanitized, @"[A-Za-z]:\\[^\s]+", "[path removed]");
            sanitized = Regex.Replace(sanitized, @"\/[^\s]+", "[path removed]");
        }
        
        // Limit message length (but allow longer messages for detailed errors)
        var maxLength = preserveDetails ? 2000 : 500;
        if (sanitized.Length > maxLength)
        {
            sanitized = sanitized.Substring(0, maxLength - 3) + "...";
        }

        return sanitized.Trim();
    }
}

