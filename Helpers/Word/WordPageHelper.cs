using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper class for Word page operations.
/// </summary>
public static class WordPageHelper
{
    /// <summary>
    ///     Gets the list of section indices to operate on based on provided parameters.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="sectionIndex">Optional single section index.</param>
    /// <param name="sectionIndices">Optional array of section indices.</param>
    /// <param name="validateRange">Whether to validate that indices are within range.</param>
    /// <returns>A list of section indices to operate on.</returns>
    /// <exception cref="ArgumentException">Thrown when section index is out of range.</exception>
    public static List<int> GetTargetSections(Document doc, int? sectionIndex, JsonArray? sectionIndices,
        bool validateRange = true)
    {
        if (sectionIndices is { Count: > 0 })
            return ParseSectionIndicesArray(doc, sectionIndices, validateRange);

        if (sectionIndex.HasValue)
            return ParseSingleSectionIndex(doc, sectionIndex.Value, validateRange);

        return Enumerable.Range(0, doc.Sections.Count).ToList();
    }

    /// <summary>
    ///     Parses section indices from a JSON array.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="sectionIndices">The JSON array of section indices.</param>
    /// <param name="validateRange">Whether to validate indices are within range.</param>
    /// <returns>A list of section indices.</returns>
    private static List<int> ParseSectionIndicesArray(Document doc, JsonArray sectionIndices, bool validateRange)
    {
        var indices = sectionIndices
            .Select(s => s?.GetValue<int>())
            .Where(s => s.HasValue)
            .Select(s => s!.Value)
            .ToList();

        if (validateRange)
            ValidateSectionIndices(doc, indices);

        return indices;
    }

    /// <summary>
    ///     Parses a single section index.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="index">The section index.</param>
    /// <param name="validateRange">Whether to validate the index is within range.</param>
    /// <returns>A list containing the single section index.</returns>
    /// <exception cref="ArgumentException">Thrown when index is out of range.</exception>
    private static List<int> ParseSingleSectionIndex(Document doc, int index, bool validateRange)
    {
        List<int> indices = [index];
        if (validateRange)
            ValidateSectionIndices(doc, indices);
        return indices;
    }

    /// <summary>
    ///     Validates that all section indices are within range.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="indices">The list of indices to validate.</param>
    /// <exception cref="ArgumentException">Thrown when any index is out of range.</exception>
    private static void ValidateSectionIndices(Document doc, List<int> indices)
    {
        foreach (var idx in indices)
            if (idx < 0 || idx >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex {idx} must be between 0 and {doc.Sections.Count - 1}");
    }
}
