using System.Reflection;

namespace AsposeMcpServer.Core.Handlers;

/// <summary>
///     Registry for operation handlers with manual registration and auto-discovery support.
///     Each Tool creates its own registry and registers relevant handlers.
/// </summary>
/// <typeparam name="TContext">The document context type.</typeparam>
/// <remarks>
///     <para>
///         This registry provides a simple way to map operation names to handlers.
///         Operation names are matched case-insensitively.
///     </para>
///     <para>
///         Auto-discovery example:
///         <code>
///         var registry = HandlerRegistry&lt;Document&gt;.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Text");
///         var handler = registry.GetHandler("add");
///         var result = handler.Execute(context, parameters);
///         </code>
///     </para>
///     <para>
///         Manual registration example:
///         <code>
///         var registry = new HandlerRegistry&lt;Document&gt;();
///         registry.Register(new AddTextHandler());
///         registry.Register(new DeleteTextHandler());
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
    ///     Creates a handler registry by auto-discovering all handlers in the specified namespace.
    /// </summary>
    /// <param name="targetNamespace">
    ///     The namespace to scan for handlers implementing <see cref="IOperationHandler{TContext}" />.
    /// </param>
    /// <param name="assembly">
    ///     The assembly to scan. If null, uses <see cref="Assembly.GetExecutingAssembly" />.
    /// </param>
    /// <returns>A configured handler registry with all discovered handlers registered.</returns>
    /// <exception cref="ArgumentNullException">
    ///     Thrown when <paramref name="targetNamespace" /> is null or empty.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when a handler cannot be instantiated or when duplicate operations are detected.
    /// </exception>
    public static HandlerRegistry<TContext> CreateFromNamespace(string targetNamespace, Assembly? assembly = null)
    {
        if (string.IsNullOrEmpty(targetNamespace))
            throw new ArgumentNullException(nameof(targetNamespace));

        var registry = new HandlerRegistry<TContext>();
        var handlerInterface = typeof(IOperationHandler<TContext>);
        var scanAssembly = assembly ?? Assembly.GetExecutingAssembly();

        var handlerTypes = scanAssembly.GetTypes()
            .Where(t => t.Namespace == targetNamespace &&
                        t.IsClass &&
                        !t.IsAbstract &&
                        handlerInterface.IsAssignableFrom(t) &&
                        !t.IsDefined(typeof(ExcludeFromAutoDiscoveryAttribute), false) &&
                        t.GetConstructor(Type.EmptyTypes) != null);

        foreach (var handlerType in handlerTypes)
            try
            {
                var handler = (IOperationHandler<TContext>)Activator.CreateInstance(handlerType)!;
                registry.Register(handler);
            }
            catch (Exception ex) when (ex is not InvalidOperationException)
            {
                throw new InvalidOperationException(
                    $"Failed to instantiate handler '{handlerType.FullName}': {ex.Message}", ex);
            }

        return registry;
    }

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
