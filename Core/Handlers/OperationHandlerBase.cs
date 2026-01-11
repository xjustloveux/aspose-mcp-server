using System.Text.Json;

namespace AsposeMcpServer.Core.Handlers;

/// <summary>
///     Base class for operation handlers providing common functionality.
///     Inherit from this class to implement specific operation handlers.
/// </summary>
/// <typeparam name="TContext">
///     The document context type (e.g., Aspose.Words.Document, Aspose.Slides.Presentation).
/// </typeparam>
/// <remarks>
///     <para>
///         This base class provides utility methods for creating result messages
///         and managing document modification state.
///     </para>
///     <para>
///         Example implementation:
///         <code>
///         public class AddTextHandler : OperationHandlerBase&lt;Document&gt;
///         {
///             public override string Operation =&gt; "add";
/// 
///             public override string Execute(OperationContext&lt;Document&gt; context, OperationParameters parameters)
///             {
///                 var text = parameters.GetRequired&lt;string&gt;("text");
///                 // ... perform operation ...
///                 MarkModified(context);
///                 return Success($"Added {text.Length} characters");
///             }
///         }
///         </code>
///     </para>
/// </remarks>
public abstract class OperationHandlerBase<TContext> : IOperationHandler<TContext>
    where TContext : class
{
    // ReSharper disable once StaticMemberInGenericType - Intentional: shared options instance across all handler types
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    /// <inheritdoc />
    public abstract string Operation { get; }

    /// <inheritdoc />
    public abstract string Execute(OperationContext<TContext> context, OperationParameters parameters);

    /// <summary>
    ///     Creates a success result message.
    /// </summary>
    /// <param name="message">The success message describing the operation outcome.</param>
    /// <returns>The formatted success message.</returns>
    protected static string Success(string message)
    {
        return message;
    }

    /// <summary>
    ///     Creates a JSON result from an object.
    ///     Uses camelCase property naming for consistency with MCP conventions.
    /// </summary>
    /// <param name="data">The data object to serialize.</param>
    /// <returns>The JSON string representation.</returns>
    protected static string JsonResult(object data)
    {
        return JsonSerializer.Serialize(data, JsonOptions);
    }

    /// <summary>
    ///     Marks the context as modified for session-based operations.
    ///     Call this method after making any changes to the document.
    ///     The Tool layer uses this flag to determine whether to save the document.
    /// </summary>
    /// <param name="context">The operation context to mark as modified.</param>
    protected static void MarkModified(OperationContext<TContext> context)
    {
        context.IsModified = true;
    }
}
