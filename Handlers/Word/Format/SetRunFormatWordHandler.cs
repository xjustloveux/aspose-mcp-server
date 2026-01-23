using System.Drawing;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Handler for setting run format in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
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
        return new SuccessResult { Message = $"Run format updated{colorMsg}" };
    }

    /// <summary>
    ///     Extracts font parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted font parameters.</returns>
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

    /// <summary>
    ///     Ensures that runs exist in the paragraph, creating one if necessary.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="para">The paragraph.</param>
    /// <param name="runIndex">The run index being targeted.</param>
    /// <returns>The list of runs in the paragraph.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraph has no runs and runIndex is not 0.</exception>
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

    /// <summary>
    ///     Gets the runs to format based on run index.
    /// </summary>
    /// <param name="runs">The list of all runs.</param>
    /// <param name="runIndex">The specific run index, or null for all runs.</param>
    /// <returns>The list of runs to format.</returns>
    /// <exception cref="ArgumentException">Thrown when run index is out of range.</exception>
    private static List<Run> GetRunsToFormat(List<Run> runs, int? runIndex)
    {
        if (!runIndex.HasValue) return runs;

        if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
            throw new ArgumentException(
                $"runIndex must be between 0 and {runs.Count - 1} (paragraph has {runs.Count} Runs)");

        return [runs[runIndex.Value]];
    }

    /// <summary>
    ///     Applies font formatting to a run.
    /// </summary>
    /// <param name="run">The run to format.</param>
    /// <param name="fontParams">The font parameters.</param>
    /// <param name="isAutoColor">Whether to use auto color.</param>
    private static void ApplyFontFormatting(Run run, FontParameters fontParams, bool isAutoColor)
    {
        var underlineStr = GetUnderlineString(fontParams.Underline);

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

    /// <summary>
    ///     Applies font settings when using auto color mode.
    /// </summary>
    /// <param name="run">The run to format.</param>
    /// <param name="fontParams">The font parameters.</param>
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

    /// <summary>
    ///     Converts nullable bool underline value to string representation.
    /// </summary>
    /// <param name="underline">The nullable underline value.</param>
    /// <returns>The underline string: "single", "none", or null.</returns>
    private static string? GetUnderlineString(bool? underline)
    {
        if (!underline.HasValue) return null;
        return underline.Value ? "single" : "none";
    }

    /// <summary>
    ///     Record to hold font formatting parameters.
    /// </summary>
    /// <param name="FontName">The font name.</param>
    /// <param name="FontNameAscii">The ASCII font name.</param>
    /// <param name="FontNameFarEast">The Far East font name.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="Bold">Whether to apply bold.</param>
    /// <param name="Italic">Whether to apply italic.</param>
    /// <param name="Underline">Whether to apply underline.</param>
    /// <param name="Color">The font color.</param>
    private sealed record FontParameters(
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
