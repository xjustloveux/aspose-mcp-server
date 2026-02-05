using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.FormField;

/// <summary>
///     Result for form data export operation.
/// </summary>
public record ExportFormDataResult
{
    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }

    /// <summary>
    ///     The format used for export (fdf, xfdf, or xml).
    /// </summary>
    [JsonPropertyName("format")]
    public required string Format { get; init; }

    /// <summary>
    ///     The path to the exported data file.
    /// </summary>
    [JsonPropertyName("exportPath")]
    public required string ExportPath { get; init; }
}
