using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Field;

/// <summary>
///     Check box form field information.
/// </summary>
public sealed record CheckBoxFormFieldInfo : FormFieldInfo
{
    /// <summary>
    ///     Indicates whether the checkbox is checked.
    /// </summary>
    [JsonPropertyName("isChecked")]
    public required bool IsChecked { get; init; }
}
