using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Handlers.Excel.OleObject;

/// <summary>
///     Handler for the <c>extract_all</c> operation on <c>excel_ole_object</c>. Iterates
///     every OLE across all sheets with cumulative-byte-cap enforcement (F-8).
/// </summary>
[ResultType(typeof(OleExtractAllResult))]
public sealed class ExtractAllExcelOleObjectHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "extract_all";

    /// <summary>Extracts all embedded OLE objects to the output directory.</summary>
    /// <param name="context">Operation context.</param>
    /// <param name="parameters">Required: <c>outputDirectory</c>.</param>
    /// <returns>An <see cref="OleExtractAllResult" /> describing writes and skips.</returns>
    /// <exception cref="ArgumentException">Thrown when the output directory fails validation.</exception>
    /// <exception cref="UnauthorizedAccessException">
    ///     Thrown when the output directory cannot be created or is not writable.
    /// </exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(parameters);

        var outputDirectory = parameters.GetRequired<string>(OleParamKeys.OutputDirectory);
        OleHandlerShared.ValidateOutputDirectory(outputDirectory, context.ServerConfig);
        OleHandlerShared.EnsureDirectoryWritable(outputDirectory);

        var cap = context.ServerConfig?.MaxExtractAllBytes ?? long.MaxValue;
        if (cap <= 0) cap = long.MaxValue;

        var resolver = new OleCollisionResolver();
        var items = new List<OleExtractAllItem>();
        var skipped = new List<OleSkippedEntry>();
        long cumulative = 0;
        var truncated = false;
        var flatIndex = 0;
        var stop = false;

        for (var si = 0; si < context.Document.Worksheets.Count && !stop; si++)
        {
            var ws = context.Document.Worksheets[si];
            foreach (var ole in ws.OleObjects)
            {
                var thisIndex = flatIndex++;
                if (ole.IsLink)
                {
                    skipped.Add(new OleSkippedEntry(thisIndex, "linked"));
                    continue;
                }

                var data = ole.ObjectData;
                if (data == null || data.Length == 0)
                {
                    skipped.Add(new OleSkippedEntry(thisIndex, "empty-payload"));
                    continue;
                }

                if (cumulative + data.Length > cap)
                {
                    skipped.Add(new OleSkippedEntry(thisIndex, "cumulative-size-cap-exceeded"));
                    truncated = true;
                    stop = true;
                    break;
                }

                var metadata = ExcelOleMetadataMapper.Map(ole, ws, si, thisIndex);
                var fileName = Path.GetFileName(resolver.Reserve(outputDirectory, metadata.SuggestedFileName));
                var outputPath = Path.Combine(outputDirectory, fileName);
                // H18: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
                outputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath,
                    context.ServerConfig?.AllowedBasePaths ?? [], nameof(outputPath));

                try
                {
                    File.WriteAllBytes(outputPath, data);
                }
                catch (Exception ex)
                {
                    throw OleErrorTranslator.Translate(ex, fileName);
                }

                cumulative += data.Length;
                items.Add(new OleExtractAllItem
                {
                    Index = thisIndex,
                    OutputFilePath = Path.GetFullPath(outputPath),
                    BytesWritten = data.Length,
                    SanitizedFromRaw = !string.Equals(metadata.SuggestedFileName, metadata.RawFileName,
                        StringComparison.Ordinal)
                });
            }
        }

        return new OleExtractAllResult
        {
            Requested = flatIndex,
            Extracted = items.Count,
            Skipped = skipped,
            Items = items,
            Truncated = truncated
        };
    }
}
