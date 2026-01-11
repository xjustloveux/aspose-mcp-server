using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Handlers.Word.Page;

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
        {
            var indices = sectionIndices
                .Select(s => s?.GetValue<int>())
                .Where(s => s.HasValue)
                .Select(s => s!.Value)
                .ToList();

            if (validateRange)
                foreach (var idx in indices)
                    if (idx < 0 || idx >= doc.Sections.Count)
                        throw new ArgumentException(
                            $"sectionIndex {idx} must be between 0 and {doc.Sections.Count - 1}");

            return indices;
        }

        if (sectionIndex.HasValue)
        {
            if (validateRange && (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count))
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            return [sectionIndex.Value];
        }

        return Enumerable.Range(0, doc.Sections.Count).ToList();
    }
}
