using System.Text.Json.Nodes;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Interface for Aspose MCP tools that provide document manipulation capabilities
/// </summary>
public interface IAsposeTool
{
    /// <summary>
    /// Gets the description of the tool and its usage examples
    /// </summary>
    string Description { get; }
    
    /// <summary>
    /// Gets the JSON schema defining the input parameters for the tool
    /// </summary>
    object InputSchema { get; }
    
    /// <summary>
    /// Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    Task<string> ExecuteAsync(JsonObject? arguments);
}

/// <summary>
/// Optional interface for tools that support MCP 2025-11-25 annotations.
/// Annotations provide metadata about tool behavior for security and UX purposes.
/// </summary>
public interface IAnnotatedTool : IAsposeTool
{
    /// <summary>
    /// Indicates if the tool is read-only (does not modify data).
    /// If null, annotation is not provided.
    /// </summary>
    bool? IsReadOnly { get; }
    
    /// <summary>
    /// Indicates if the tool is destructive (may delete or permanently modify data).
    /// If null, annotation is not provided.
    /// </summary>
    bool? IsDestructive { get; }
}

