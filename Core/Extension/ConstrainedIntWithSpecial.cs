namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Integer setting with special value support (e.g., 0 = never unload).
/// </summary>
public class ConstrainedIntWithSpecial : ConstrainedInt
{
    /// <summary>
    ///     Initializes a new instance with default values.
    /// </summary>
    public ConstrainedIntWithSpecial()
    {
    }

    /// <summary>
    ///     Initializes a new instance with specified values.
    /// </summary>
    /// <param name="defaultValue">Default value when extension doesn't specify.</param>
    /// <param name="floor">Minimum allowed value for normal values.</param>
    /// <param name="ceiling">Maximum allowed value for normal values.</param>
    /// <param name="specialValue">The special value with magic meaning.</param>
    /// <param name="specialAllowed">Whether extensions can use the special value.</param>
    public ConstrainedIntWithSpecial(int defaultValue, int floor, int ceiling,
        int specialValue, bool specialAllowed = true)
        : base(defaultValue, floor, ceiling)
    {
        SpecialValue = specialValue;
        SpecialAllowed = specialAllowed;
    }

    /// <summary>
    ///     The special value with magic meaning (e.g., 0 for "never unload").
    /// </summary>
    public int SpecialValue { get; set; }

    /// <summary>
    ///     Whether extensions are allowed to use the special value.
    /// </summary>
    public bool SpecialAllowed { get; set; } = true;

    /// <summary>
    ///     Applies constraint with special value handling.
    /// </summary>
    /// <param name="extensionValue">The value specified by extension, or null to use default.</param>
    /// <returns>The constrained value, or special value if allowed.</returns>
    public new int Apply(int? extensionValue)
    {
        if (extensionValue == SpecialValue)
            return SpecialAllowed ? SpecialValue : Default;

        return base.Apply(extensionValue);
    }

    /// <summary>
    ///     Applies constraint with special value handling and returns warning if constrained.
    /// </summary>
    /// <param name="extensionValue">The value specified by extension, or null to use default.</param>
    /// <param name="settingName">Name of the setting for warning message.</param>
    /// <returns>Constraint result with value and optional warning message.</returns>
    public new ConstraintResult ApplyWithWarning(int? extensionValue, string settingName)
    {
        if (extensionValue == SpecialValue)
        {
            if (SpecialAllowed)
                return new ConstraintResult(SpecialValue, null);

            return new ConstraintResult(Default,
                $"{settingName}={SpecialValue} (special value) is not allowed, using default {Default}");
        }

        return base.ApplyWithWarning(extensionValue, settingName);
    }
}
