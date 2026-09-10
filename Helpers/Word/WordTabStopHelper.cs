using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Reads the <c>tabStops</c> array once, for every operation that takes one.
///     <para>
///         R13-W01: four handlers read that array and each read it differently. A null entry was
///         silently skipped by one and turned into a stop at position 0 by the other three. A
///         missing <c>position</c> meant 0 everywhere, which is the left margin and therefore no
///         stop at all. A misspelt alignment or leader was silently taken as left/none by three of
///         them, and <c>add_styled</c> matched its names case-sensitively in PascalCase while the
///         rest matched lower case — so <c>"center"</c> was honoured by three handlers and
///         quietly read as left by the fourth. One documented parameter, five behaviours.
///     </para>
///     <para>
///         The resolve-then-apply order matters as much as the parsing. Reading values inside the
///         apply loop meant a malformed entry threw after <c>TabStops.Clear()</c>, leaving the
///         paragraph with no stops at all and the request reporting failure (§23.4). That was
///         fixed in <c>edit</c> and left in place in the other three, which is the same shape as
///         every other "the rule was not on the path that needed it" defect in this codebase: the
///         fix has to live where every caller reaches it, so it lives here.
///     </para>
/// </summary>
public static class WordTabStopHelper
{
    /// <summary>
    ///     Most stops one paragraph may be given.
    /// </summary>
    /// <remarks>
    ///     Measured against Aspose.Words 23.10 rather than chosen: 1,500 stops on one paragraph
    ///     survive a save and reload unchanged, while 2,000 come back as 1,947. Above the bound
    ///     the request would be accepted and then silently not be what was stored, which is the
    ///     outcome a cap exists to prevent.
    /// </remarks>
    public const int MaxTabStops = 1500;

    /// <summary>
    ///     Furthest a stop may be placed, in points, in either direction.
    /// </summary>
    /// <remarks>
    ///     22 inches, the largest page Word will lay out; a stop beyond it can never be reached.
    ///     Nothing rejects a larger number further down — measured, a position of 1e9 is accepted
    ///     and reads back as 107,374,182.35, the point at which the underlying twip field
    ///     saturates — so it is rejected here instead of being silently rewritten.
    /// </remarks>
    public const double MaxPositionPoints = 1584;

    /// <summary>
    ///     Turns the requested tab stops into the objects to apply, refusing anything it cannot
    ///     read. Nothing is mutated: the caller clears and applies only once this has returned.
    /// </summary>
    /// <param name="tabStops">The tab stops array, or null.</param>
    /// <returns>The resolved stops, in the order given.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when there are more stops than <see cref="MaxTabStops" />, when an entry is not
    ///     an object, when its <c>position</c> is absent, not a number, not finite or beyond
    ///     <see cref="MaxPositionPoints" />, or when its <c>alignment</c> or <c>leader</c> is not
    ///     one this understands.
    /// </exception>
    public static List<TabStop> Resolve(JsonArray? tabStops)
    {
        List<TabStop> resolved = [];
        if (tabStops is not { Count: > 0 }) return resolved;

        if (tabStops.Count > MaxTabStops)
            throw new ArgumentException(
                $"'tabStops' holds {tabStops.Count} entries, more than the {MaxTabStops} a "
                + "paragraph stores without losing some. Measured: 1,500 survive a save and "
                + "reload unchanged, 2,000 come back as 1,947.");

        for (var index = 0; index < tabStops.Count; index++)
        {
            if (tabStops[index] is not JsonObject entry)
                throw new ArgumentException(
                    $"Tab stop {index}: expected an object with a 'position', but was "
                    + $"{tabStops[index]?.ToJsonString() ?? "null"}.");

            resolved.Add(new TabStop(
                PositionOf(entry, index),
                WordParagraphHelper.GetTabAlignment(TextOf(entry, "alignment", "left", index)),
                WordParagraphHelper.GetTabLeader(TextOf(entry, "leader", "none", index))));
        }

        return resolved;
    }

    /// <summary>Reads an entry's position, in points.</summary>
    /// <param name="entry">The tab stop entry.</param>
    /// <param name="index">Its index, for the message.</param>
    /// <returns>The position.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the position is absent, not a number, not finite, or out of range. Absent
    ///     used to mean 0, which places the stop on the left margin — a request that asked for a
    ///     stop and got none, reported as success (R13-W01).
    /// </exception>
    private static double PositionOf(JsonObject entry, int index)
    {
        var node = entry["position"];
        if (node == null)
            throw new ArgumentException(
                $"Tab stop {index}: 'position' is required, in points from the left margin.");

        double position;
        try
        {
            position = node.GetValue<double>();
        }
        catch (Exception ex) when (ex is InvalidOperationException or FormatException)
        {
            throw new ArgumentException(
                $"Tab stop {index}: 'position' must be a number, but was {node.ToJsonString()}.");
        }

        if (double.IsNaN(position) || double.IsInfinity(position))
            throw new ArgumentException($"Tab stop {index}: 'position' must be a finite number.");

        if (Math.Abs(position) > MaxPositionPoints)
            throw new ArgumentException(
                $"Tab stop {index}: 'position' must be between -{MaxPositionPoints} and "
                + $"{MaxPositionPoints} points; {position} is off any page Word will lay out.");

        return position;
    }

    /// <summary>Reads one of an entry's string fields.</summary>
    /// <param name="entry">The tab stop entry.</param>
    /// <param name="name">The field to read.</param>
    /// <param name="fallback">What an absent field means.</param>
    /// <param name="index">The entry's index, for the message.</param>
    /// <returns>The field's value, or the fallback.</returns>
    /// <exception cref="ArgumentException">Thrown when the field is present but not a string.</exception>
    private static string TextOf(JsonObject entry, string name, string fallback, int index)
    {
        var node = entry[name];
        if (node == null) return fallback;

        try
        {
            return node.GetValue<string>();
        }
        catch (Exception ex) when (ex is InvalidOperationException or FormatException)
        {
            throw new ArgumentException(
                $"Tab stop {index}: '{name}' must be a string, but was {node.ToJsonString()}.");
        }
    }
}
