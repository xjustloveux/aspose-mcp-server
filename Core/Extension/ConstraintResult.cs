namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Result of applying a constraint to a value.
/// </summary>
/// <param name="Value">The constrained value.</param>
/// <param name="Warning">Warning message if the value was constrained, null otherwise.</param>
public record ConstraintResult(int Value, string? Warning)
{
    /// <summary>
    ///     Gets a value indicating whether the value was constrained and a warning was generated.
    /// </summary>
    public bool HasWarning => Warning is not null;
}
