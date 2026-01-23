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
///         This base class provides utility methods for managing document modification state.
///     </para>
///     <para>
///         Example implementation:
///         <code>
///         [ResultType(typeof(SuccessResult))]
///         public class AddTextHandler : OperationHandlerBase&lt;Document&gt;
///         {
///             public override string Operation =&gt; "add";
/// 
///             public override object Execute(OperationContext&lt;Document&gt; context, OperationParameters parameters)
///             {
///                 var text = parameters.GetRequired&lt;string&gt;("text");
///                 // ... perform operation ...
///                 MarkModified(context);
///                 return new SuccessResult { Message = $"Added {text.Length} characters" };
///             }
///         }
///         </code>
///     </para>
/// </remarks>
public abstract class OperationHandlerBase<TContext> : IOperationHandler<TContext>
    where TContext : class
{
    /// <inheritdoc />
    public abstract string Operation { get; }

    /// <inheritdoc />
    public abstract object Execute(OperationContext<TContext> context, OperationParameters parameters);

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
