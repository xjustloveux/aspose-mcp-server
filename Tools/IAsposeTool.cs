using System.Text.Json.Nodes;

namespace AsposeMcpServer.Tools;

public interface IAsposeTool
{
    string Description { get; }
    object InputSchema { get; }
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

