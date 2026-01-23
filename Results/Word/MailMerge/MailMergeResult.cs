using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.MailMerge;

/// <summary>
///     Result for mail merge operations on Word documents.
/// </summary>
public record MailMergeResult
{
    /// <summary>
    ///     Description of the template source (file path or session ID).
    /// </summary>
    [JsonPropertyName("templateSource")]
    public required string TemplateSource { get; init; }

    /// <summary>
    ///     Number of fields merged.
    /// </summary>
    [JsonPropertyName("fieldsMerged")]
    public required int FieldsMerged { get; init; }

    /// <summary>
    ///     Number of records processed (1 for single record, more for multiple records).
    /// </summary>
    [JsonPropertyName("recordsProcessed")]
    public required int RecordsProcessed { get; init; }

    /// <summary>
    ///     Cleanup options that were applied.
    /// </summary>
    [JsonPropertyName("cleanupApplied")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CleanupApplied { get; init; }

    /// <summary>
    ///     List of output files created.
    /// </summary>
    [JsonPropertyName("outputFiles")]
    public required IReadOnlyList<string> OutputFiles { get; init; }
}
