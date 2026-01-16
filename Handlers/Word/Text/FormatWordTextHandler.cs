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
        var p = ExtractFormatParameters(parameters);

        var doc = context.Document;

        ValidateSectionIndex(doc, p.SectionIndex);
        var section = doc.Sections[p.SectionIndex];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        ValidateParagraphIndex(paragraphs, p.ParagraphIndex, p.SectionIndex);
        var para = paragraphs[p.ParagraphIndex];

        var runs = EnsureRuns(doc, para, p.ParagraphIndex);
        var runsToFormat = GetRunsToFormat(runs, p.RunIndex);

        var changes = ApplyFormatting(runsToFormat, p);

        MarkModified(context);

        return BuildResultMessage(p.ParagraphIndex, p.RunIndex, runsToFormat.Count, changes);
    }

    /// <summary>
    ///     Extracts format parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted format parameters.</returns>
    private static FormatParameters ExtractFormatParameters(OperationParameters parameters)
    {
        return new FormatParameters(
            parameters.GetRequired<int>("paragraphIndex"),
            parameters.GetOptional<int?>("runIndex"),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<string?>("fontNameAscii"),
            parameters.GetOptional<string?>("fontNameFarEast"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional<bool?>("italic"),
            parameters.GetOptional<string?>("underline"),
            parameters.GetOptional<string?>("color"),
            parameters.GetOptional<bool?>("strikethrough"),
            parameters.GetOptional<bool?>("superscript"),
            parameters.GetOptional<bool?>("subscript"));
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
    /// <param name="p">The format parameters containing font settings.</param>
    /// <returns>List of applied changes.</returns>
    private static List<string> ApplyFormatting(List<Run> runsToFormat, FormatParameters p)
    {
        List<string> changes = [];

        foreach (var run in runsToFormat)
        {
            HandleSuperscriptSubscriptExclusivity(run, p.Superscript, p.Subscript);

            FontHelper.Word.ApplyFontSettings(
                run, p.FontName, p.FontNameAscii, p.FontNameFarEast, p.FontSize,
                p.Bold, p.Italic, p.Underline, p.Color, p.Strikethrough, p.Superscript, p.Subscript);

            CollectChanges(changes, p);
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
    /// <param name="changes">The list to add changes to.</param>
    /// <param name="p">The format parameters containing font settings.</param>
    private static void CollectChanges(List<string> changes, FormatParameters p)
    {
        if (!string.IsNullOrEmpty(p.FontNameAscii))
            changes.Add($"Font (ASCII): {p.FontNameAscii}");
        if (!string.IsNullOrEmpty(p.FontNameFarEast))
            changes.Add($"Font (Far East): {p.FontNameFarEast}");
        if (!string.IsNullOrEmpty(p.FontName) && string.IsNullOrEmpty(p.FontNameAscii) &&
            string.IsNullOrEmpty(p.FontNameFarEast))
            changes.Add($"Font: {p.FontName}");
        if (p.FontSize.HasValue)
            changes.Add($"Font size: {p.FontSize.Value} points");
        if (p.Bold.HasValue)
            changes.Add($"Bold: {(p.Bold.Value ? "Yes" : "No")}");
        if (p.Italic.HasValue)
            changes.Add($"Italic: {(p.Italic.Value ? "Yes" : "No")}");
        if (!string.IsNullOrEmpty(p.Underline))
            changes.Add($"Underline: {p.Underline}");
        if (!string.IsNullOrEmpty(p.Color))
        {
            var colorValue = p.Color.TrimStart('#');
            changes.Add($"Color: {(colorValue.Length == 6 ? "#" : "")}{colorValue}");
        }

        if (p.Strikethrough.HasValue)
            changes.Add($"Strikethrough: {(p.Strikethrough.Value ? "Yes" : "No")}");
        if (p.Superscript.HasValue)
            changes.Add($"Superscript: {(p.Superscript.Value ? "Yes" : "No")}");
        if (p.Subscript.HasValue)
            changes.Add($"Subscript: {(p.Subscript.Value ? "Yes" : "No")}");
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

    /// <summary>
    ///     Record to hold format parameters.
    /// </summary>
    /// <param name="ParagraphIndex">The paragraph index.</param>
    /// <param name="RunIndex">The run index.</param>
    /// <param name="SectionIndex">The section index.</param>
    /// <param name="FontName">The font name.</param>
    /// <param name="FontNameAscii">The ASCII font name.</param>
    /// <param name="FontNameFarEast">The Far East font name.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="Bold">Whether to apply bold.</param>
    /// <param name="Italic">Whether to apply italic.</param>
    /// <param name="Underline">The underline style.</param>
    /// <param name="Color">The font color.</param>
    /// <param name="Strikethrough">Whether to apply strikethrough.</param>
    /// <param name="Superscript">Whether to apply superscript.</param>
    /// <param name="Subscript">Whether to apply subscript.</param>
    private sealed record FormatParameters(
        int ParagraphIndex,
        int? RunIndex,
        int SectionIndex,
        string? FontName,
        string? FontNameAscii,
        string? FontNameFarEast,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        string? Underline,
        string? Color,
        bool? Strikethrough,
        bool? Superscript,
        bool? Subscript);
}
