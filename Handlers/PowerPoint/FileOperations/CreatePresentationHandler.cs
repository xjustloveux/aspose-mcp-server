using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.FileOperations;

/// <summary>
///     Handler for creating a new PowerPoint presentation.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class CreatePresentationHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "create";

    /// <summary>
    ///     Creates a new PowerPoint presentation.
    /// </summary>
    /// <param name="context">The presentation context (not used for create).</param>
    /// <param name="parameters">
    ///     Required: path or outputPath
    /// </param>
    /// <returns>Success message with output path.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        // Held for the whole operation: this handler builds or reads presentations of
        // its own, outside any session, so nothing else stands between it and a library
        // that fails when two threads are inside it (SlidesGate).
        using var slidesGate = SlidesGate.Enter();

        var p = ExtractCreateParameters(parameters);

        var savePath = p.Path ?? p.OutputPath;
        if (string.IsNullOrEmpty(savePath))
            throw new ArgumentException("path or outputPath is required for create operation");

        SecurityHelper.ValidateFilePath(savePath, allowAbsolutePaths: true);

        using var presentation = new Presentation();
        // H21: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
        savePath = SecurityHelper.ResolveAndEnsureWithinAllowlist(savePath,
            context.ServerConfig?.AllowedBasePaths ?? [], nameof(savePath));
        presentation.Save(savePath, PptSaveFormatResolver.Resolve(savePath));

        return new SuccessResult { Message = $"PowerPoint presentation created successfully. Output: {savePath}" };
    }

    /// <summary>
    ///     Extracts create parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted create parameters.</returns>
    private static CreateParameters ExtractCreateParameters(OperationParameters parameters)
    {
        return new CreateParameters(
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("outputPath"));
    }

    /// <summary>
    ///     Record for holding create presentation parameters.
    /// </summary>
    /// <param name="Path">The output file path.</param>
    /// <param name="OutputPath">Alternative output file path.</param>
    private sealed record CreateParameters(string? Path, string? OutputPath);
}
