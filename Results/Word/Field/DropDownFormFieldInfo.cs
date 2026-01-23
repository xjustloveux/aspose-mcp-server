using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Field;

/// <summary>
///     Drop-down form field information.
/// </summary>
public sealed record DropDownFormFieldInfo : FormFieldInfo
{
    /// <summary>
    ///     Index of the selected item.
    /// </summary>
    [JsonPropertyName("selectedIndex")]
    public required int SelectedIndex { get; init; }

    /// <summary>
    ///     List of available options.
    /// </summary>
    [JsonPropertyName("options")]
    public required IReadOnlyList<string> Options { get; init; }
}
