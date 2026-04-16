using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Handlers.Excel.OleObject;

/// <summary>
///     Handler for the <c>extract</c> operation on <c>excel_ole_object</c>.
/// </summary>
[ResultType(typeof(OleExtractResult))]
public sealed class ExtractExcelOleObjectHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "extract";

    /// <summary>Extracts a single OLE object by flat index.</summary>
    /// <param name="context">Operation context.</param>
    /// <param name="parameters">Required: <c>oleIndex</c>, <c>outputDirectory</c>. Optional: <c>outputFileName</c>.</param>
    /// <returns>An <see cref="OleExtractResult" /> with the absolute output path.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the index is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the OLE object is linked.</exception>
    /// <exception cref="IOException">Thrown when writing fails.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(parameters);

        var oleIndex = parameters.GetRequired<int>(OleParamKeys.OleIndex);
        var outputDirectory = parameters.GetRequired<string>(OleParamKeys.OutputDirectory);
        var overrideName = parameters.GetOptional<string?>(OleParamKeys.OutputFileName);

        OleHandlerShared.ValidateOutputDirectory(outputDirectory, context.ServerConfig);

        var (ole, sheet, sheetIndex, _) = OleHandlerShared.LocateExcelOle(context.Document, oleIndex);
        if (ole.IsLink) throw new InvalidOperationException(OleErrorMessageBuilder.LinkedCannotExtract());
        var data = ole.ObjectData;
        if (data == null || data.Length == 0) throw new IOException(OleErrorMessageBuilder.SaveFailed(null));

        var metadata = ExcelOleMetadataMapper.Map(ole, sheet, sheetIndex, oleIndex);
        var fileName = OleHandlerShared.ResolveExtractFileName(metadata, overrideName);
        OleHandlerShared.EnsureDirectoryWritable(outputDirectory);
        var outputPath = Path.Combine(outputDirectory, fileName);
        // H17: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
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

        return new OleExtractResult
        {
            Index = oleIndex,
            OutputFilePath = Path.GetFullPath(outputPath),
            BytesWritten = data.Length,
            SanitizedFromRaw =
                !string.Equals(metadata.SuggestedFileName, metadata.RawFileName, StringComparison.Ordinal)
        };
    }
}
