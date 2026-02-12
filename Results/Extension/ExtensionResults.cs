namespace AsposeMcpServer.Results.Extension;

/// <summary>
///     Static class containing all extension result types for schema generation.
/// </summary>
public static class ExtensionResults
{
    /// <summary>
    ///     All result types used by ExtensionTool.
    /// </summary>
    public static readonly Type[] AllTypes =
    [
        typeof(ListExtensionsResult),
        typeof(ExtensionInfoDto),
        typeof(BindExtensionResult),
        typeof(UnbindExtensionResult),
        typeof(SetFormatResult),
        typeof(ExtensionStatusResult),
        typeof(ExtensionBindingsResult),
        typeof(BindingInfoDto)
    ];
}
