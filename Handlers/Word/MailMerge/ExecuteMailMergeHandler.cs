using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.MailMerging;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.MailMerge;

/// <summary>
///     Handler for executing mail merge operations on Word documents.
/// </summary>
public class ExecuteMailMergeHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "execute";

    /// <summary>
    ///     Executes mail merge on a Word document template.
    /// </summary>
    /// <param name="context">The document context containing the template.</param>
    /// <param name="parameters">
    ///     Required: outputPath, and either data or dataArray.
    ///     Optional: cleanupOptions.
    /// </param>
    /// <returns>Success message with field and file information.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when required parameters are missing or both data and dataArray are provided.
    /// </exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractMailMergeParameters(parameters);

        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);

        JsonObject? dataObject = null;
        JsonArray? dataArrayObject = null;

        if (!string.IsNullOrEmpty(p.Data))
            dataObject = JsonNode.Parse(p.Data) as JsonObject;

        if (!string.IsNullOrEmpty(p.DataArray))
            dataArrayObject = JsonNode.Parse(p.DataArray) as JsonArray;

        if (dataObject == null && dataArrayObject == null)
            throw new ArgumentException(
                "Either 'data' (for single record) or 'dataArray' (for multiple records) must be provided");

        if (dataObject != null && dataArrayObject != null)
            throw new ArgumentException(
                "Cannot specify both 'data' and 'dataArray'. Use 'data' for single record or 'dataArray' for multiple records");

        var cleanupOptionsFlags = ParseCleanupOptions(p.CleanupOptions);

        if (dataArrayObject is { Count: > 0 })
            return ExecuteMultipleRecords(context, p.OutputPath, dataArrayObject, cleanupOptionsFlags);

        if (dataObject != null)
            return ExecuteSingleRecord(context, p.OutputPath, dataObject, cleanupOptionsFlags);

        throw new ArgumentException("No data provided for mail merge");
    }

    private static MailMergeParameters ExtractMailMergeParameters(OperationParameters parameters)
    {
        return new MailMergeParameters(
            parameters.GetRequired<string>("outputPath"),
            parameters.GetOptional<string?>("data"),
            parameters.GetOptional<string?>("dataArray"),
            parameters.GetOptional<string?>("cleanupOptions"));
    }

    /// <summary>
    ///     Executes mail merge for a single record.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="data">The JSON object containing field names and values.</param>
    /// <param name="cleanupOptions">The mail merge cleanup options to apply.</param>
    /// <returns>A message indicating the result of the mail merge operation.</returns>
    private static string ExecuteSingleRecord(OperationContext<Document> context, string outputPath, JsonObject data,
        MailMergeCleanupOptions cleanupOptions)
    {
        var doc = context.Document.Clone() ?? throw new InvalidOperationException("Failed to clone document");

        doc.MailMerge.CleanupOptions = cleanupOptions;

        var fieldNames = data.Select(kvp => kvp.Key).ToArray();
        var fieldValues = data.Select(kvp => kvp.Value?.ToString() ?? "").Cast<object>().ToArray();

        doc.MailMerge.Execute(fieldNames, fieldValues);
        doc.Save(outputPath);

        var result = new StringBuilder();
        result.AppendLine("Mail merge completed successfully");
        result.AppendLine($"Template: {GetTemplateSource(context)}");
        result.AppendLine($"Output: {outputPath}");
        result.AppendLine($"Fields merged: {fieldNames.Length}");
        if (cleanupOptions != MailMergeCleanupOptions.None)
            result.AppendLine($"Cleanup applied: {cleanupOptions}");

        return Success(result.ToString());
    }

    /// <summary>
    ///     Executes mail merge for multiple records.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="outputPath">The base output file path (files will be numbered).</param>
    /// <param name="dataArray">The JSON array containing multiple record objects.</param>
    /// <param name="cleanupOptions">The mail merge cleanup options to apply.</param>
    /// <returns>A message indicating the result of the mail merge operation.</returns>
    private static string ExecuteMultipleRecords(OperationContext<Document> context, string outputPath,
        JsonArray dataArray,
        MailMergeCleanupOptions cleanupOptions)
    {
        List<string> outputFiles = [];
        var outputDir = Path.GetDirectoryName(outputPath) ?? ".";
        var outputName = Path.GetFileNameWithoutExtension(outputPath);
        var outputExt = Path.GetExtension(outputPath);

        for (var i = 0; i < dataArray.Count; i++)
        {
            var recordData = dataArray[i] as JsonObject;
            if (recordData == null) continue;

            var doc = context.Document.Clone() ?? throw new InvalidOperationException("Failed to clone document");

            doc.MailMerge.CleanupOptions = cleanupOptions;

            var fieldNames = recordData.Select(kvp => kvp.Key).ToArray();
            var fieldValues = recordData.Select(kvp => kvp.Value?.ToString() ?? "").Cast<object>().ToArray();

            doc.MailMerge.Execute(fieldNames, fieldValues);

            var recordOutputPath = dataArray.Count == 1
                ? outputPath
                : Path.Combine(outputDir, $"{outputName}_{i + 1}{outputExt}");

            doc.Save(recordOutputPath);
            outputFiles.Add(recordOutputPath);
        }

        var result = new StringBuilder();
        result.AppendLine("Mail merge completed successfully (multiple records)");
        result.AppendLine($"Template: {GetTemplateSource(context)}");
        result.AppendLine($"Records processed: {outputFiles.Count}");
        if (cleanupOptions != MailMergeCleanupOptions.None)
            result.AppendLine($"Cleanup applied: {cleanupOptions}");
        result.AppendLine("Output files:");
        foreach (var file in outputFiles) result.AppendLine($"  - {file}");

        return Success(result.ToString());
    }

    /// <summary>
    ///     Gets the template source description.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <returns>A string describing the template source.</returns>
    private static string GetTemplateSource(OperationContext<Document> context)
    {
        if (!string.IsNullOrEmpty(context.SessionId))
            return $"session {context.SessionId}";

        if (!string.IsNullOrEmpty(context.SourcePath))
            return Path.GetFileName(context.SourcePath);

        return "document";
    }

    /// <summary>
    ///     Parses cleanup options from comma-separated string.
    /// </summary>
    /// <param name="optionsString">The comma-separated cleanup options string.</param>
    /// <returns>The parsed MailMergeCleanupOptions flags.</returns>
    private static MailMergeCleanupOptions ParseCleanupOptions(string? optionsString)
    {
        if (string.IsNullOrEmpty(optionsString))
            return MailMergeCleanupOptions.RemoveUnusedFields | MailMergeCleanupOptions.RemoveEmptyParagraphs;

        var options = MailMergeCleanupOptions.None;
        var optionsList =
            optionsString.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var option in optionsList)
            options |= option.ToLower() switch
            {
                "removeunusedfields" => MailMergeCleanupOptions.RemoveUnusedFields,
                "removeunusedregions" => MailMergeCleanupOptions.RemoveUnusedRegions,
                "removeemptyparagraphs" => MailMergeCleanupOptions.RemoveEmptyParagraphs,
                "removecontainingfields" => MailMergeCleanupOptions.RemoveContainingFields,
                "removestaticfields" => MailMergeCleanupOptions.RemoveStaticFields,
                _ => MailMergeCleanupOptions.None
            };

        return options;
    }

    private sealed record MailMergeParameters(
        string OutputPath,
        string? Data,
        string? DataArray,
        string? CleanupOptions);
}
