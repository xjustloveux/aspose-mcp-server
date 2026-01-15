namespace AsposeMcpServer.Core.Handlers;

/// <summary>
///     Indicates that a handler class should be excluded from auto-discovery by
///     <see cref="HandlerRegistry{TContext}.CreateFromNamespace" />.
/// </summary>
/// <remarks>
///     <para>
///         Apply this attribute to handler classes that should not be automatically
///         registered, such as base classes, test handlers, or deprecated handlers.
///     </para>
/// </remarks>
[AttributeUsage(AttributeTargets.Class, Inherited = false)]
public sealed class ExcludeFromAutoDiscoveryAttribute : Attribute;
