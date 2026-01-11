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
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontNameAscii = parameters.GetOptional<string?>("fontNameAscii");
        var fontNameFarEast = parameters.GetOptional<string?>("fontNameFarEast");
        var fontSize = parameters.GetOptional<double?>("fontSize");
        var bold = parameters.GetOptional<bool?>("bold");
        var italic = parameters.GetOptional<bool?>("italic");
        var underline = parameters.GetOptional<bool?>("underline");
        var color = parameters.GetOptional<string?>("color");

        var doc = context.Document;
        var para = WordFormatHelper.GetTargetParagraph(doc, paragraphIndex);
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();

        if (runs.Count == 0 && runIndex.HasValue)
        {
            if (runIndex.Value != 0)
                throw new ArgumentException("Paragraph has no Run nodes, runIndex must be 0 to create a new Run");
            var newRun = new Run(doc);
            para.AppendChild(newRun);
            runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        }
        else if (runs.Count == 0)
        {
            var newRun = new Run(doc);
            para.AppendChild(newRun);
            runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        }

        List<Run> runsToFormat;
        if (runIndex.HasValue)
        {
            if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
                throw new ArgumentException(
                    $"runIndex must be between 0 and {runs.Count - 1} (paragraph has {runs.Count} Runs)");
            runsToFormat = [runs[runIndex.Value]];
        }
        else
        {
            runsToFormat = runs;
        }

        var isAutoColor = color?.Equals("auto", StringComparison.OrdinalIgnoreCase) == true;

        foreach (var run in runsToFormat)
        {
            var underlineStr = underline.HasValue ? underline.Value ? "single" : "none" : null;

            if (isAutoColor)
                run.Font.Color = Color.Empty;
            else
                FontHelper.Word.ApplyFontSettings(
                    run,
                    fontName,
                    fontNameAscii,
                    fontNameFarEast,
                    fontSize,
                    bold,
                    italic,
                    underlineStr,
                    color
                );

            // Apply other font settings when auto color is set
            if (isAutoColor)
            {
                if (!string.IsNullOrEmpty(fontName)) run.Font.Name = fontName;
                if (!string.IsNullOrEmpty(fontNameAscii)) run.Font.NameAscii = fontNameAscii;
                if (!string.IsNullOrEmpty(fontNameFarEast)) run.Font.NameFarEast = fontNameFarEast;
                if (fontSize.HasValue) run.Font.Size = fontSize.Value;
                if (bold.HasValue) run.Font.Bold = bold.Value;
                if (italic.HasValue) run.Font.Italic = italic.Value;
                if (underline.HasValue) run.Font.Underline = underline.Value ? Underline.Single : Underline.None;
            }
        }

        MarkModified(context);
        var colorMsg = isAutoColor ? " (color reset to auto)" : "";
        return Success($"Run format updated{colorMsg}");
    }
}
