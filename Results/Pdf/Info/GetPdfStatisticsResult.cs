using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Info;

/// <summary>
///     Result for getting statistics from PDF documents.
/// </summary>
public record GetPdfStatisticsResult
{
    /// <summary>
    ///     File size in bytes.
    /// </summary>
    [JsonPropertyName("fileSizeBytes")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? FileSizeBytes { get; init; }

    /// <summary>
    ///     File size in kilobytes.
    /// </summary>
    [JsonPropertyName("fileSizeKb")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? FileSizeKb { get; init; }

    /// <summary>
    ///     Total number of pages in the document.
    /// </summary>
    [JsonPropertyName("totalPages")]
    public required int TotalPages { get; init; }

    /// <summary>
    ///     Whether the document is encrypted.
    /// </summary>
    [JsonPropertyName("isEncrypted")]
    public required bool IsEncrypted { get; init; }

    /// <summary>
    ///     Whether the document is linearized.
    /// </summary>
    [JsonPropertyName("isLinearized")]
    public required bool IsLinearized { get; init; }

    /// <summary>
    ///     Number of bookmarks in the document.
    /// </summary>
    [JsonPropertyName("bookmarks")]
    public required int Bookmarks { get; init; }

    /// <summary>
    ///     Number of form fields in the document.
    /// </summary>
    [JsonPropertyName("formFields")]
    public required int FormFields { get; init; }

    /// <summary>
    ///     Total number of annotations in the document.
    /// </summary>
    [JsonPropertyName("totalAnnotations")]
    public required int TotalAnnotations { get; init; }

    /// <summary>
    ///     Total number of paragraphs in the document.
    /// </summary>
    [JsonPropertyName("totalParagraphs")]
    public required int TotalParagraphs { get; init; }

    /// <summary>
    ///     Note about session mode limitations.
    /// </summary>
    [JsonPropertyName("note")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; init; }
}
