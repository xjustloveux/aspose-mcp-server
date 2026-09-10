using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.MailMerging;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Word.MailMerge;

namespace AsposeMcpServer.Handlers.Word.MailMerge;

/// <summary>
///     Handler for executing mail merge operations on Word documents.
/// </summary>
[ResultType(typeof(MailMergeResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractMailMergeParameters(parameters);

        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);
        // H3: resolve symlinks once at handler entry; the resolved path is forwarded to the private helpers.
        var resolvedOutputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputPath,
            context.ServerConfig?.AllowedBasePaths ?? [], "outputPath");

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

        // The records arrive as a JSON string, so the bound applied to array parameters at the
        // tool boundary never saw them: the count only exists after decoding. Each record clones
        // the whole document and writes a file, so both the count and what it writes are bounded
        // before the first clone rather than after some of them exist (R3-R05).
        if (dataArrayObject != null)
        {
            SecurityHelper.ValidateArraySize(dataArrayObject, "dataArray");
            RenderBudget.EnsureOutputCount(dataArrayObject.Count, "merged documents");
        }

        var cleanupOptionsFlags = ParseCleanupOptions(p.CleanupOptions);

        if (dataArrayObject is { Count: > 0 })
            return ExecuteMultipleRecords(context, resolvedOutputPath, dataArrayObject, cleanupOptionsFlags);

        if (dataObject != null)
            return ExecuteSingleRecord(context, resolvedOutputPath, dataObject, cleanupOptionsFlags);

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
    /// <returns>A result containing the mail merge operation details.</returns>
    private static MailMergeResult ExecuteSingleRecord(OperationContext<Document> context, string outputPath,
        JsonObject data,
        MailMergeCleanupOptions cleanupOptions)
    {
        var doc = context.Document.Clone() ?? throw new InvalidOperationException("Failed to clone document");

        doc.MailMerge.CleanupOptions = cleanupOptions;

        var templateFieldNames = doc.MailMerge.GetFieldNames();
        var fieldNames = data.Select(kvp => kvp.Key).ToArray();
        var fieldValues = data.Select(kvp => kvp.Value?.ToString() ?? "").Cast<object>().ToArray();

        doc.MailMerge.Execute(fieldNames, fieldValues);

        // The single-record path wrote straight to the destination with no size limit at all,
        // while the multi-record path measured only after the file was already there (R4-R02).
        // Both now produce the file elsewhere and publish it once it is complete and within
        // budget, so a refusal leaves the caller's destination as it was.
        BoundedFilePublisher.Publish(outputPath, RenderBudget.MaxOutputBytes,
            stream => doc.Save(stream, SaveFormat.Docx), "merged document",
            context.Recovery, context.ServerConfig?.AllowedBasePaths ?? []);

        var actualMergedCount = fieldNames.Count(f => templateFieldNames.Contains(f));

        return new MailMergeResult
        {
            TemplateSource = GetTemplateSource(context),
            FieldsMerged = actualMergedCount,
            RecordsProcessed = 1,
            CleanupApplied = cleanupOptions != MailMergeCleanupOptions.None ? cleanupOptions.ToString() : null,
            OutputFiles = [outputPath]
        };
    }

    /// <summary>
    ///     Executes mail merge for multiple records.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="outputPath">The base output file path (files will be numbered).</param>
    /// <param name="dataArray">The JSON array containing multiple record objects.</param>
    /// <param name="cleanupOptions">The mail merge cleanup options to apply.</param>
    /// <returns>A result containing the mail merge operation details.</returns>
    private static MailMergeResult ExecuteMultipleRecords(OperationContext<Document> context, string outputPath,
        JsonArray dataArray,
        MailMergeCleanupOptions cleanupOptions)
    {
        var outputDir = Path.GetDirectoryName(outputPath) ?? ".";
        var outputName = Path.GetFileNameWithoutExtension(outputPath);
        var outputExt = Path.GetExtension(outputPath);
        var fieldsMerged = 0;

        // Every record is staged and the batch published at the end. Publishing each record as it
        // was produced meant a merge refused on its last record had already replaced the
        // destinations of every record before it, leaving a caller who retried unable to tell
        // which files belonged to the run that failed (R4-R02).
        using var batch = new BoundedFileBatch(RenderBudget.MaxOutputBytes, "merged documents",
            context.Recovery, context.ServerConfig?.AllowedBasePaths ?? []);

        for (var i = 0; i < dataArray.Count; i++)
        {
            var recordData = dataArray[i] as JsonObject;
            if (recordData == null) continue;

            var doc = context.Document.Clone() ?? throw new InvalidOperationException("Failed to clone document");

            doc.MailMerge.CleanupOptions = cleanupOptions;

            var templateFieldNames = doc.MailMerge.GetFieldNames();
            var fieldNames = recordData.Select(kvp => kvp.Key).ToArray();
            var fieldValues = recordData.Select(kvp => kvp.Value?.ToString() ?? "").Cast<object>().ToArray();

            doc.MailMerge.Execute(fieldNames, fieldValues);
            var actualMergedCount = fieldNames.Count(f => templateFieldNames.Contains(f));
            fieldsMerged = Math.Max(fieldsMerged, actualMergedCount);

            var recordOutputPath = dataArray.Count == 1
                ? outputPath
                : Path.Combine(outputDir, $"{outputName}_{i + 1}{outputExt}");
            // H3: re-resolve each per-record path immediately before its sink (bug 20260415-symlink-toctou-sweep).
            recordOutputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(recordOutputPath,
                context.ServerConfig?.AllowedBasePaths ?? [], nameof(recordOutputPath));
            // A record count says nothing about what lands on disk; one template can produce a
            // very large document, so each is written through the remaining budget rather than
            // published first and measured afterwards (R3-R05, R4-R02).
            batch.Stage(recordOutputPath, stream => doc.Save(stream, SaveFormat.Docx));
        }

        var outputFiles = batch.Publish();

        return new MailMergeResult
        {
            TemplateSource = GetTemplateSource(context),
            FieldsMerged = fieldsMerged,
            RecordsProcessed = outputFiles.Count,
            CleanupApplied = cleanupOptions != MailMergeCleanupOptions.None ? cleanupOptions.ToString() : null,
            OutputFiles = outputFiles
        };
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
