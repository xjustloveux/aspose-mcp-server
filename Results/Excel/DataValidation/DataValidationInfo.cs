using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.DataValidation;

/// <summary>
///     Information about a single data validation.
/// </summary>
public record DataValidationInfo
{
    /// <summary>
    ///     Zero-based index of the data validation.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Validation type.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Operator type.
    /// </summary>
    [JsonPropertyName("operatorType")]
    public required string OperatorType { get; init; }

    /// <summary>
    ///     First formula.
    /// </summary>
    [JsonPropertyName("formula1")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula1 { get; init; }

    /// <summary>
    ///     Second formula.
    /// </summary>
    [JsonPropertyName("formula2")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula2 { get; init; }

    /// <summary>
    ///     Error message.
    /// </summary>
    [JsonPropertyName("errorMessage")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ErrorMessage { get; init; }

    /// <summary>
    ///     Input message.
    /// </summary>
    [JsonPropertyName("inputMessage")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? InputMessage { get; init; }

    /// <summary>
    ///     Whether to show error.
    /// </summary>
    [JsonPropertyName("showError")]
    public required bool ShowError { get; init; }

    /// <summary>
    ///     Whether to show input.
    /// </summary>
    [JsonPropertyName("showInput")]
    public required bool ShowInput { get; init; }

    /// <summary>
    ///     Whether to show in-cell drop down.
    /// </summary>
    [JsonPropertyName("inCellDropDown")]
    public required bool InCellDropDown { get; init; }
}
