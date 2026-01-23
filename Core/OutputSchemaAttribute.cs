namespace AsposeMcpServer.Core;

/// <summary>
///     Specifies the type to use for generating the OutputSchema of an MCP tool.
///     This allows tools that return 'object' to still have proper schema generation.
/// </summary>
[AttributeUsage(AttributeTargets.Method, Inherited = false)]
public class OutputSchemaAttribute : Attribute
{
    /// <summary>
    ///     Initializes a new instance of the OutputSchemaAttribute class.
    /// </summary>
    /// <param name="schemaType">The type to use for generating the output schema.</param>
    public OutputSchemaAttribute(Type schemaType)
    {
        SchemaType = schemaType ?? throw new ArgumentNullException(nameof(schemaType));
    }

    /// <summary>
    ///     Gets the type to use for schema generation.
    /// </summary>
    public Type SchemaType { get; }
}
