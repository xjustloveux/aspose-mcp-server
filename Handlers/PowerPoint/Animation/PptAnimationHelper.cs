using Aspose.Slides.Animation;

namespace AsposeMcpServer.Handlers.PowerPoint.Animation;

/// <summary>
///     Helper class for PowerPoint animation operations.
/// </summary>
public static class PptAnimationHelper
{
    /// <summary>
    ///     Parses effect type string to EffectType enum.
    /// </summary>
    /// <param name="effectTypeStr">The effect type string to parse.</param>
    /// <returns>The parsed EffectType enum value, or Fade as default.</returns>
    public static EffectType ParseEffectType(string? effectTypeStr)
    {
        if (string.IsNullOrEmpty(effectTypeStr)) return EffectType.Fade;
        return Enum.TryParse<EffectType>(effectTypeStr, true, out var result) ? result : EffectType.Fade;
    }

    /// <summary>
    ///     Parses effect subtype string to EffectSubtype enum.
    /// </summary>
    /// <param name="subtypeStr">The effect subtype string to parse.</param>
    /// <returns>The parsed EffectSubtype enum value, or None as default.</returns>
    public static EffectSubtype ParseEffectSubtype(string? subtypeStr)
    {
        if (string.IsNullOrEmpty(subtypeStr)) return EffectSubtype.None;
        return Enum.TryParse<EffectSubtype>(subtypeStr, true, out var result) ? result : EffectSubtype.None;
    }

    /// <summary>
    ///     Parses trigger type string to EffectTriggerType enum.
    /// </summary>
    /// <param name="triggerTypeStr">The trigger type string to parse.</param>
    /// <returns>The parsed EffectTriggerType enum value, or OnClick as default.</returns>
    public static EffectTriggerType ParseTriggerType(string? triggerTypeStr)
    {
        if (string.IsNullOrEmpty(triggerTypeStr)) return EffectTriggerType.OnClick;
        return Enum.TryParse<EffectTriggerType>(triggerTypeStr, true, out var result)
            ? result
            : EffectTriggerType.OnClick;
    }
}
