namespace AsposeMcpServer.Core;

/// <summary>
///     Specifies the Handler namespace for a Tool, used for collecting result types to generate outputSchema.
/// </summary>
[AttributeUsage(AttributeTargets.Class, Inherited = false)]
public class ToolHandlerMappingAttribute : Attribute
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ToolHandlerMappingAttribute" /> class.
    /// </summary>
    /// <param name="handlerNamespace">The namespace containing the Handler classes.</param>
    /// <exception cref="ArgumentNullException">Thrown when handlerNamespace is null.</exception>
    public ToolHandlerMappingAttribute(string handlerNamespace)
    {
        HandlerNamespace = handlerNamespace ?? throw new ArgumentNullException(nameof(handlerNamespace));
    }

    /// <summary>
    ///     Gets the Handler namespace for result type collection.
    /// </summary>
    public string HandlerNamespace { get; }
}
