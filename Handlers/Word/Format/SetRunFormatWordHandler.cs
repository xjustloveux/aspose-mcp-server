using System.Drawing;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Handler for setting run format in Word documents.
/// </summary>
public class SetRunFormatWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_run_format";

    /// <summary>
    ///     Sets run format properties.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex
    ///     Optional: runIndex, fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic, underline, color
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetOptional("paragraphIndex", 0);
        var runIndex = parameters.GetOptional<int?>("runIndex");
        var fontParams = ExtractFontParameters(parameters);

        var doc = context.Document;
        var para = WordFormatHelper.GetTargetParagraph(doc, paragraphIndex);
        var runs = EnsureRunsExist(doc, para, runIndex);
        var runsToFormat = GetRunsToFormat(runs, runIndex);

        var isAutoColor = fontParams.Color?.Equals("auto", StringComparison.OrdinalIgnoreCase) == true;

        foreach (var run in runsToFormat) ApplyFontFormatting(run, fontParams, isAutoColor);

        MarkModified(context);
        var colorMsg = isAutoColor ? " (color reset to auto)" : "";
        return Success($"Run format updated{colorMsg}");
    }

    private static FontParameters ExtractFontParameters(OperationParameters parameters)
    {
        return new FontParameters(
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<string?>("fontNameAscii"),
            parameters.GetOptional<string?>("fontNameFarEast"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional<bool?>("italic"),
            parameters.GetOptional<bool?>("underline"),
            parameters.GetOptional<string?>("color")
        );
    }

    private static List<Run> EnsureRunsExist(Document doc, Aspose.Words.Paragraph para, int? runIndex)
    {
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();

        if (runs.Count > 0) return runs;

        if (runIndex.HasValue && runIndex.Value != 0)
            throw new ArgumentException("Paragraph has no Run nodes, runIndex must be 0 to create a new Run");

        var newRun = new Run(doc);
        para.AppendChild(newRun);
        return para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
    }

    private static List<Run> GetRunsToFormat(List<Run> runs, int? runIndex)
    {
        if (!runIndex.HasValue) return runs;

        if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
            throw new ArgumentException(
                $"runIndex must be between 0 and {runs.Count - 1} (paragraph has {runs.Count} Runs)");

        return [runs[runIndex.Value]];
    }

    private static void ApplyFontFormatting(Run run, FontParameters fontParams, bool isAutoColor)
    {
        var underlineStr = fontParams.Underline.HasValue ? fontParams.Underline.Value ? "single" : "none" : null;

        if (isAutoColor)
        {
            run.Font.Color = Color.Empty;
            ApplyAutoColorFontSettings(run, fontParams);
        }
        else
        {
            FontHelper.Word.ApplyFontSettings(
                run,
                fontParams.FontName,
                fontParams.FontNameAscii,
                fontParams.FontNameFarEast,
                fontParams.FontSize,
                fontParams.Bold,
                fontParams.Italic,
                underlineStr,
                fontParams.Color
            );
        }
    }

    private static void ApplyAutoColorFontSettings(Run run, FontParameters fontParams)
    {
        if (!string.IsNullOrEmpty(fontParams.FontName)) run.Font.Name = fontParams.FontName;
        if (!string.IsNullOrEmpty(fontParams.FontNameAscii)) run.Font.NameAscii = fontParams.FontNameAscii;
        if (!string.IsNullOrEmpty(fontParams.FontNameFarEast)) run.Font.NameFarEast = fontParams.FontNameFarEast;
        if (fontParams.FontSize.HasValue) run.Font.Size = fontParams.FontSize.Value;
        if (fontParams.Bold.HasValue) run.Font.Bold = fontParams.Bold.Value;
        if (fontParams.Italic.HasValue) run.Font.Italic = fontParams.Italic.Value;
        if (fontParams.Underline.HasValue)
            run.Font.Underline = fontParams.Underline.Value ? Underline.Single : Underline.None;
    }

    private record FontParameters(
        string? FontName,
        string? FontNameAscii,
        string? FontNameFarEast,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        bool? Underline,
        string? Color
    );
}
