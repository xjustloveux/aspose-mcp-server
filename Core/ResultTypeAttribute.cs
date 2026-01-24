namespace AsposeMcpServer.Core;

/// <summary>
///     Specifies the result type for a Handler, used for generating outputSchema.
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class ResultTypeAttribute : Attribute
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ResultTypeAttribute" /> class.
    /// </summary>
    /// <param name="resultType">The type to use for generating the output schema.</param>
    /// <exception cref="ArgumentNullException">Thrown when resultType is null.</exception>
    public ResultTypeAttribute(Type resultType)
    {
        ResultType = resultType ?? throw new ArgumentNullException(nameof(resultType));
    }

    /// <summary>
    ///     Gets the type to use for schema generation.
    /// </summary>
    public Type ResultType { get; }
}
