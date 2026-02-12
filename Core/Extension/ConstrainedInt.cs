namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Integer setting with default value and range constraint.
/// </summary>
public class ConstrainedInt
{
    /// <summary>
    ///     Initializes a new instance with default values.
    /// </summary>
    // ReSharper disable once MemberCanBeProtected.Global - Required public for JSON deserialization
    public ConstrainedInt()
    {
    }

    /// <summary>
    ///     Initializes a new instance with specified values.
    /// </summary>
    /// <param name="defaultValue">Default value when extension doesn't specify.</param>
    /// <param name="floor">Minimum allowed value.</param>
    /// <param name="ceiling">Maximum allowed value.</param>
    public ConstrainedInt(int defaultValue, int floor, int ceiling)
    {
        Default = defaultValue;
        Floor = floor;
        Ceiling = ceiling;
    }

    /// <summary>
    ///     Default value when extension doesn't specify.
    /// </summary>
    public int Default { get; set; }

    /// <summary>
    ///     Minimum allowed value (floor). Extensions cannot go below this.
    /// </summary>
    public int Floor { get; set; }

    /// <summary>
    ///     Maximum allowed value (ceiling). Extensions cannot go above this.
    /// </summary>
    public int Ceiling { get; set; }

    /// <summary>
    ///     Applies constraint to extension value.
    /// </summary>
    /// <param name="extensionValue">The value specified by extension, or null to use default.</param>
    /// <returns>The constrained value within [Floor, Ceiling] range.</returns>
    public int Apply(int? extensionValue)
    {
        return Math.Clamp(extensionValue ?? Default, Floor, Ceiling);
    }

    /// <summary>
    ///     Applies constraint and returns warning if value was constrained.
    /// </summary>
    /// <param name="extensionValue">The value specified by extension, or null to use default.</param>
    /// <param name="settingName">Name of the setting for warning message.</param>
    /// <returns>Constraint result with value and optional warning message.</returns>
    // ReSharper disable once MemberCanBeProtected.Global - Public API for direct use on ConstrainedInt
    public ConstraintResult ApplyWithWarning(int? extensionValue, string settingName)
    {
        var value = extensionValue ?? Default;
        var result = Math.Clamp(value, Floor, Ceiling);

        if (extensionValue.HasValue && result != extensionValue.Value)
        {
            var direction = extensionValue.Value < Floor ? "below minimum" : "above maximum";
            return new ConstraintResult(result,
                $"{settingName}={extensionValue.Value} is {direction}, constrained to {result}");
        }

        return new ConstraintResult(result, null);
    }
}
