using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Handler for formatting existing text in Word documents.
/// </summary>
public class FormatWordTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "format";

    /// <summary>
    ///     Formats text at a specific paragraph and run position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex.
    ///     Optional: runIndex, sectionIndex, fontName, fontNameAscii, fontNameFarEast,
    ///     fontSize, bold, italic, underline, color, strikethrough, superscript, subscript.
    /// </param>
    /// <returns>Success message with formatting details.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraphIndex is missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the paragraph has no runs.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetRequired<int>("paragraphIndex");
        var runIndex = parameters.GetOptional<int?>("runIndex");
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontNameAscii = parameters.GetOptional<string?>("fontNameAscii");
        var fontNameFarEast = parameters.GetOptional<string?>("fontNameFarEast");
        var fontSize = parameters.GetOptional<double?>("fontSize");
        var bold = parameters.GetOptional<bool?>("bold");
        var italic = parameters.GetOptional<bool?>("italic");
        var underline = parameters.GetOptional<string?>("underline");
        var color = parameters.GetOptional<string?>("color");
        var strikethrough = parameters.GetOptional<bool?>("strikethrough");
        var superscript = parameters.GetOptional<bool?>("superscript");
        var subscript = parameters.GetOptional<bool?>("subscript");

        var doc = context.Document;

        ValidateSectionIndex(doc, sectionIndex);
        var section = doc.Sections[sectionIndex];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        ValidateParagraphIndex(paragraphs, paragraphIndex, sectionIndex);
        var para = paragraphs[paragraphIndex];

        var runs = EnsureRuns(doc, para, paragraphIndex);
        var runsToFormat = GetRunsToFormat(runs, runIndex);

        var changes = ApplyFormatting(runsToFormat, fontName, fontNameAscii, fontNameFarEast,
            fontSize, bold, italic, underline, color, strikethrough, superscript, subscript);

        MarkModified(context);

        return BuildResultMessage(paragraphIndex, runIndex, runsToFormat.Count, changes);
    }

    /// <summary>
    ///     Validates the section index is within range.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="sectionIndex">The section index to validate.</param>
    /// <exception cref="ArgumentException">Thrown when section index is out of range.</exception>
    private static void ValidateSectionIndex(Document doc, int sectionIndex)
    {
        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
            throw new ArgumentException(
                $"sectionIndex {sectionIndex} is out of range (document has {doc.Sections.Count} sections)");
    }

    /// <summary>
    ///     Validates the paragraph index is within range.
    /// </summary>
    /// <param name="paragraphs">The list of paragraphs.</param>
    /// <param name="paragraphIndex">The paragraph index to validate.</param>
    /// <param name="sectionIndex">The section index for error message.</param>
    /// <exception cref="ArgumentException">Thrown when paragraph index is out of range.</exception>
    private static void ValidateParagraphIndex(List<WordParagraph> paragraphs, int paragraphIndex, int sectionIndex)
    {
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex} is out of range (section {sectionIndex} body has {paragraphs.Count} paragraphs)");
    }

    /// <summary>
    ///     Ensures the paragraph has runs, creating one if necessary.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="para">The paragraph.</param>
    /// <param name="paragraphIndex">The paragraph index for error message.</param>
    /// <returns>The collection of runs.</returns>
    /// <exception cref="InvalidOperationException">Thrown when runs cannot be created.</exception>
    private static NodeCollection EnsureRuns(Document doc, WordParagraph? para, int paragraphIndex)
    {
        if (para == null)
            throw new ArgumentNullException(nameof(para), "Paragraph cannot be null");

        var runs = para.GetChildNodes(NodeType.Run, false);
        if (runs == null || runs.Count == 0)
        {
            var newRun = new Run(doc, "");
            para.AppendChild(newRun);
            runs = para.GetChildNodes(NodeType.Run, false);
            if (runs == null || runs.Count == 0)
                throw new InvalidOperationException(
                    $"Paragraph #{paragraphIndex} has no Run nodes and cannot create new Run node");
        }

        return runs;
    }

    /// <summary>
    ///     Gets the runs to format based on run index.
    /// </summary>
    /// <param name="runs">All runs in the paragraph.</param>
    /// <param name="runIndex">Optional specific run index.</param>
    /// <returns>List of runs to format.</returns>
    /// <exception cref="ArgumentException">Thrown when run index is out of range.</exception>
    private static List<Run> GetRunsToFormat(NodeCollection runs, int? runIndex)
    {
        List<Run> runsToFormat = [];

        if (runIndex.HasValue)
        {
            if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
                throw new ArgumentException(
                    $"Run index {runIndex.Value} is out of range (paragraph has {runs.Count} Runs)");
            if (runs[runIndex.Value] is Run run)
                runsToFormat.Add(run);
        }
        else
        {
            foreach (var node in runs)
                if (node is Run run)
                    runsToFormat.Add(run);
        }

        return runsToFormat;
    }

    /// <summary>
    ///     Applies formatting to the specified runs.
    /// </summary>
    /// <param name="runsToFormat">The runs to format.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size.</param>
    /// <param name="bold">Whether text should be bold.</param>
    /// <param name="italic">Whether text should be italic.</param>
    /// <param name="underline">The underline style.</param>
    /// <param name="color">The text color.</param>
    /// <param name="strikethrough">Whether text should have strikethrough.</param>
    /// <param name="superscript">Whether text should be superscript.</param>
    /// <param name="subscript">Whether text should be subscript.</param>
    /// <returns>List of applied changes.</returns>
    private static List<string> ApplyFormatting(List<Run> runsToFormat, string? fontName, string? fontNameAscii,
        string? fontNameFarEast, double? fontSize, bool? bold, bool? italic, string? underline, string? color,
        bool? strikethrough, bool? superscript, bool? subscript)
    {
        List<string> changes = [];

        foreach (var run in runsToFormat)
        {
            HandleSuperscriptSubscriptExclusivity(run, superscript, subscript);

            FontHelper.Word.ApplyFontSettings(
                run, fontName, fontNameAscii, fontNameFarEast, fontSize,
                bold, italic, underline, color, strikethrough, superscript, subscript);

            CollectChanges(changes, fontName, fontNameAscii, fontNameFarEast, fontSize,
                bold, italic, underline, color, strikethrough, superscript, subscript);
        }

        return changes;
    }

    /// <summary>
    ///     Handles the mutual exclusivity of superscript and subscript.
    /// </summary>
    /// <param name="run">The run to modify.</param>
    /// <param name="superscript">Whether to set superscript.</param>
    /// <param name="subscript">Whether to set subscript.</param>
    private static void HandleSuperscriptSubscriptExclusivity(Run run, bool? superscript, bool? subscript)
    {
        if (!superscript.HasValue && !subscript.HasValue) return;

        if (superscript is true)
        {
            run.Font.Subscript = false;
        }
        else if (subscript is true)
        {
            run.Font.Superscript = false;
        }
        else
        {
            if (superscript is false) run.Font.Superscript = false;
            if (subscript is false) run.Font.Subscript = false;
        }
    }

    /// <summary>
    ///     Collects the list of formatting changes for the result message.
    /// </summary>
    private static void CollectChanges(List<string> changes, string? fontName, string? fontNameAscii,
        string? fontNameFarEast, double? fontSize, bool? bold, bool? italic, string? underline, string? color,
        bool? strikethrough, bool? superscript, bool? subscript)
    {
        if (!string.IsNullOrEmpty(fontNameAscii))
            changes.Add($"Font (ASCII): {fontNameAscii}");
        if (!string.IsNullOrEmpty(fontNameFarEast))
            changes.Add($"Font (Far East): {fontNameFarEast}");
        if (!string.IsNullOrEmpty(fontName) && string.IsNullOrEmpty(fontNameAscii) &&
            string.IsNullOrEmpty(fontNameFarEast))
            changes.Add($"Font: {fontName}");
        if (fontSize.HasValue)
            changes.Add($"Font size: {fontSize.Value} points");
        if (bold.HasValue)
            changes.Add($"Bold: {(bold.Value ? "Yes" : "No")}");
        if (italic.HasValue)
            changes.Add($"Italic: {(italic.Value ? "Yes" : "No")}");
        if (!string.IsNullOrEmpty(underline))
            changes.Add($"Underline: {underline}");
        if (!string.IsNullOrEmpty(color))
        {
            var colorValue = color.TrimStart('#');
            changes.Add($"Color: {(colorValue.Length == 6 ? "#" : "")}{colorValue}");
        }

        if (strikethrough.HasValue)
            changes.Add($"Strikethrough: {(strikethrough.Value ? "Yes" : "No")}");
        if (superscript.HasValue)
            changes.Add($"Superscript: {(superscript.Value ? "Yes" : "No")}");
        if (subscript.HasValue)
            changes.Add($"Subscript: {(subscript.Value ? "Yes" : "No")}");
    }

    /// <summary>
    ///     Builds the result message.
    /// </summary>
    /// <param name="paragraphIndex">The paragraph index.</param>
    /// <param name="runIndex">The run index if specified.</param>
    /// <param name="formattedCount">The number of runs formatted.</param>
    /// <param name="changes">The list of changes made.</param>
    /// <returns>The formatted result message.</returns>
    private static string BuildResultMessage(int paragraphIndex, int? runIndex, int formattedCount,
        List<string> changes)
    {
        var result = "Run-level formatting set successfully.";
        result += $" Paragraph index: {paragraphIndex}.";
        result +=
            runIndex.HasValue ? $" Run index: {runIndex.Value}." : $" Formatted Runs: {formattedCount}.";
        result += changes.Count > 0
            ? $" Changes: {string.Join(", ", changes.Distinct())}."
            : " No change parameters provided.";

        return Success(result);
    }
}
