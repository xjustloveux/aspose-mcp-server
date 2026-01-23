using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Field;

/// <summary>
///     Base information for a form field.
/// </summary>
[JsonDerivedType(typeof(TextFormFieldInfo), "text")]
[JsonDerivedType(typeof(CheckBoxFormFieldInfo), "checkbox")]
[JsonDerivedType(typeof(DropDownFormFieldInfo), "dropdown")]
public record FormFieldInfo
{
    /// <summary>
    ///     Zero-based index of the form field.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Name of the form field.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Form field type (FieldFormTextInput, FieldFormCheckBox, FieldFormDropDown).
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }
}
