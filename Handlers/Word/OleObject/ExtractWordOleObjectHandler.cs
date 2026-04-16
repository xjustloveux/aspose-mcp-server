using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Handlers.Word.OleObject;

/// <summary>
///     Handler for the <c>extract</c> operation on <c>word_ole_object</c>. Writes a
///     single OLE payload to disk, chosen by zero-based flat index.
/// </summary>
[ResultType(typeof(OleExtractResult))]
public sealed class ExtractWordOleObjectHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "extract";

    /// <summary>
    ///     Extracts one OLE object by index to the validated output directory.
    /// </summary>
    /// <param name="context">Operation context; <c>Document</c> must be non-null.</param>
    /// <param name="parameters">
    ///     Required: <c>oleIndex</c>, <c>outputDirectory</c>. Optional: <c>outputFileName</c>.
    /// </param>
    /// <returns>An <see cref="OleExtractResult" /> with the absolute output path.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when <c>oleIndex</c> is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the OLE object is linked.</exception>
    /// <exception cref="ArgumentException">Thrown when the output directory fails validation.</exception>
    /// <exception cref="UnauthorizedAccessException">
    ///     Thrown when the output directory cannot be created or is not writable.
    /// </exception>
    /// <exception cref="IOException">
    ///     Thrown when writing extracted bytes fails, or when the embedded payload is
    ///     zero bytes (cross-tool parity with Excel / PowerPoint extract, AC-19).
    /// </exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(parameters);

        var oleIndex = parameters.GetRequired<int>(OleParamKeys.OleIndex);
        var outputDirectory = parameters.GetRequired<string>(OleParamKeys.OutputDirectory);
        var outputFileNameOverride = parameters.GetOptional<string?>(OleParamKeys.OutputFileName);

        OleHandlerShared.ValidateOutputDirectory(outputDirectory, context.ServerConfig);

        var (shape, flatIndex) = OleHandlerShared.LocateWordShape(context.Document, oleIndex);
        var ole = shape.OleFormat!;
        if (ole.IsLink) throw new InvalidOperationException(OleErrorMessageBuilder.LinkedCannotExtract());

        var size = WordOleMetadataMapper.ComputeSize(shape);
        if (size == 0) throw new IOException(OleErrorMessageBuilder.SaveFailed(null));
        var location = WordOleMetadataMapper.ResolveLocation(context.Document, shape);
        var metadata = WordOleMetadataMapper.Map(shape, flatIndex, location, size);

        var fileName = OleHandlerShared.ResolveExtractFileName(metadata, outputFileNameOverride);
        OleHandlerShared.EnsureDirectoryWritable(outputDirectory);
        var outputPath = Path.Combine(outputDirectory, fileName);
        // H9: resolve symlinks immediately before the FileStream sink (bug 20260415-symlink-toctou-sweep).
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

        return new OleExtractResult
        {
            Index = flatIndex,
            OutputFilePath = Path.GetFullPath(outputPath),
            BytesWritten = written,
            SanitizedFromRaw =
                !string.Equals(metadata.SuggestedFileName, metadata.RawFileName, StringComparison.Ordinal)
        };
    }
}
