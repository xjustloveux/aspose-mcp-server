using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Handlers.Word.OleObject;

/// <summary>
///     Handler for the <c>extract_all</c> operation on <c>word_ole_object</c>. Iterates
///     every OLE-bearing shape and writes each embedded payload to
///     <c>outputDirectory</c>, skipping linked objects and empty payloads with explicit
///     entries in <see cref="OleExtractAllResult.Skipped" />. Enforces the F-8 cumulative
///     byte cap.
/// </summary>
[ResultType(typeof(OleExtractAllResult))]
public sealed class ExtractAllWordOleObjectHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "extract_all";

    /// <summary>
    ///     Extracts every embedded OLE payload in the document.
    /// </summary>
    /// <param name="context">Operation context; <c>Document</c> must be non-null.</param>
    /// <param name="parameters">Required: <c>outputDirectory</c>.</param>
    /// <returns>An <see cref="OleExtractAllResult" /> summarizing successes and skips.</returns>
    /// <exception cref="ArgumentException">Thrown when the output directory fails validation.</exception>
    /// <exception cref="UnauthorizedAccessException">
    ///     Thrown when the output directory cannot be created or is not writable.
    /// </exception>
    /// <exception cref="IOException">Thrown when writing any payload fails.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
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
        var stopIteration = false;
        foreach (var shape in context.Document.GetChildNodes(NodeType.Shape, true).OfType<Aspose.Words.Drawing.Shape>())
        {
            if (stopIteration) break;
            if (shape.OleFormat == null) continue;
            var ole = shape.OleFormat;
            var thisIndex = flatIndex++;

            if (ole.IsLink)
            {
                skipped.Add(new OleSkippedEntry(thisIndex, "linked"));
                continue;
            }

            var size = WordOleMetadataMapper.ComputeSize(shape);
            if (size == 0)
            {
                skipped.Add(new OleSkippedEntry(thisIndex, "empty-payload"));
                continue;
            }

            if (cumulative + size > cap)
            {
                skipped.Add(new OleSkippedEntry(thisIndex, "cumulative-size-cap-exceeded"));
                truncated = true;
                stopIteration = true;
                continue;
            }

            var location = WordOleMetadataMapper.ResolveLocation(context.Document, shape);
            var metadata = WordOleMetadataMapper.Map(shape, thisIndex, location, size);
            var fileName = Path.GetFileName(resolver.Reserve(outputDirectory, metadata.SuggestedFileName));
            var outputPath = Path.Combine(outputDirectory, fileName);
            // H10: resolve symlinks immediately before the FileStream sink (bug 20260415-symlink-toctou-sweep).
            outputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath,
                context.ServerConfig?.AllowedBasePaths ?? [], nameof(outputPath));

            long written;
            try
            {
                using var fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
                ole.Save(fs);
                fs.Flush();
                written = fs.Length;
            }
            catch (Exception ex)
            {
                throw OleErrorTranslator.Translate(ex, fileName);
            }

            cumulative += written;
            items.Add(new OleExtractAllItem
            {
                Index = thisIndex,
                OutputFilePath = Path.GetFullPath(outputPath),
                BytesWritten = written,
                SanitizedFromRaw = !string.Equals(metadata.SuggestedFileName, metadata.RawFileName,
                    StringComparison.Ordinal)
            });
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
