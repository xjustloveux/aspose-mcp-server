namespace AsposeMcpServer.Core.Handlers;

/// <summary>
///     Registry for operation handlers with manual registration support.
///     Each Tool creates its own registry and registers relevant handlers.
/// </summary>
/// <typeparam name="TContext">The document context type.</typeparam>
/// <remarks>
///     <para>
///         This registry provides a simple way to map operation names to handlers.
///         Operation names are matched case-insensitively.
///     </para>
///     <para>
///         Example usage:
///         <code>
///         var registry = new HandlerRegistry&lt;Document&gt;();
///         registry.Register(new AddTextHandler());
///         registry.Register(new DeleteTextHandler());
/// 
///         var handler = registry.GetHandler("add");
///         var result = handler.Execute(context, parameters);
///         </code>
///     </para>
/// </remarks>
public class HandlerRegistry<TContext> where TContext : class
{
    private readonly Dictionary<string, IOperationHandler<TContext>> _handlers =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    ///     Gets the count of registered handlers.
    /// </summary>
    public int Count => _handlers.Count;

    /// <summary>
    ///     Registers a handler for an operation.
    /// </summary>
    /// <param name="handler">The handler to register.</param>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when a handler for the operation is already registered.
    /// </exception>
    public void Register(IOperationHandler<TContext> handler)
    {
        if (!_handlers.TryAdd(handler.Operation, handler))
            throw new InvalidOperationException(
                $"Handler for operation '{handler.Operation}' is already registered");
    }

    /// <summary>
    ///     Gets the handler for an operation.
    /// </summary>
    /// <param name="operation">The operation name (case-insensitive).</param>
    /// <returns>The handler for the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when no handler is registered for the operation.
    ///     The exception message includes a list of available operations.
    /// </exception>
    public IOperationHandler<TContext> GetHandler(string operation)
    {
        if (!_handlers.TryGetValue(operation, out var handler))
            throw new ArgumentException(
                $"Unknown operation: {operation}. Available operations: {string.Join(", ", _handlers.Keys)}");

        return handler;
    }

    /// <summary>
    ///     Tries to get the handler for an operation without throwing.
    /// </summary>
    /// <param name="operation">The operation name (case-insensitive).</param>
    /// <param name="handler">The handler if found, otherwise null.</param>
    /// <returns>True if the handler was found, false otherwise.</returns>
    public bool TryGetHandler(string operation, out IOperationHandler<TContext>? handler)
    {
        return _handlers.TryGetValue(operation, out handler);
    }

    /// <summary>
    ///     Gets all registered operation names.
    /// </summary>
    /// <returns>Collection of registered operation names.</returns>
    public IEnumerable<string> GetOperations()
    {
        return _handlers.Keys;
    }

    /// <summary>
    ///     Checks if a handler is registered for the specified operation.
    /// </summary>
    /// <param name="operation">The operation name (case-insensitive).</param>
    /// <returns>True if a handler is registered, false otherwise.</returns>
    public bool HasHandler(string operation)
    {
        return _handlers.ContainsKey(operation);
    }
}
