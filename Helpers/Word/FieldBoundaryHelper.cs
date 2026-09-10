using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helpers for locating Word field node boundaries so text and field mutations do not corrupt
///     or nest inside existing fields. Within a paragraph a field is a flat node range:
///     FieldStart -> field-code runs -> FieldSeparator -> result runs -> FieldEnd.
/// </summary>
public static class FieldBoundaryHelper
{
    /// <summary>
    ///     Returns the innermost field whose node range (FieldStart..FieldEnd) contains the
    ///     specified node, or null when the node is not inside any field.
    ///     <para>
    ///         Answered by document position rather than by a sibling walk within the node's own
    ///         paragraph. The walk could not cross a paragraph break, so a field whose markers sit
    ///         in different paragraphs did not contain anything after the break as far as this
    ///         helper was concerned — and text insertion, deletion and replacement all use it to
    ///         decide what they may touch, so those positions were editable inside a field
    ///         (R9-W02). Taking the first match in the paragraph's own field list also returned
    ///         the outer field of a nested pair, which is the wrong answer for the same reason.
    ///     </para>
    ///     <para>
    ///         Repeated calls over one document should build a <see cref="FieldExtents" /> once
    ///         and ask it; this overload builds one per call, which numbers the whole document.
    ///     </para>
    /// </summary>
    /// <param name="node">The node to test, typically a Run.</param>
    /// <returns>The innermost enclosing field, or null.</returns>
    public static Field? GetEnclosingField(Node node)
    {
        return node.Document is Document document
            ? FieldExtents.Of(document).EnclosingField(node)
            : null;
    }

    /// <summary>
    ///     Moves the builder cursor to immediately after the field's end so a subsequent insertion
    ///     becomes a sibling following the field rather than landing inside its range.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="field">The field to move past.</param>
    public static void MoveToAfterField(DocumentBuilder builder, Field field)
    {
        if (field.End == null)
            return;

        builder.MoveTo(field.End.NextSibling ?? field.End.ParentNode);
    }

    /// <summary>
    ///     Determines whether one field lies inside another's node range.
    ///     <para>
    ///         A field's code can contain another field, and updating the outer one resolves what
    ///         is inside it. Knowing that before any update is what lets a refusal be raised before
    ///         the document has been touched rather than part-way through the traversal (R5-T01).
    ///     </para>
    /// </summary>
    /// <para>
    ///     This overload used to answer with the same sibling walk, so it disagreed with the
    ///     document-order version on exactly the case that matters: a field whose markers sit in
    ///     different paragraphs. Nothing in the server called it any more, but a helper that gives
    ///     the wrong answer is a defect waiting for its first caller, so it now numbers the
    ///     document and delegates (R8-W03).
    /// </para>
    /// <param name="inner">The field that might be nested.</param>
    /// <param name="outer">The field that might contain it.</param>
    /// <returns>True when <paramref name="inner" /> lies within <paramref name="outer" />.</returns>
    public static bool IsFieldWithinField(Field inner, Field outer)
    {
        if (ReferenceEquals(inner, outer) || inner.Start == null) return false;
        if (inner.Start.Document is not Document document) return false;

        return IsFieldWithinField(inner, outer, DocumentOrder(document));
    }

    /// <summary>
    ///     Determines whether one field lies inside another, by document position.
    /// </summary>
    /// <param name="inner">The field that might be nested.</param>
    /// <param name="outer">The field that might contain it.</param>
    /// <param name="order">Positions from <see cref="DocumentOrder" />.</param>
    /// <returns>True when <paramref name="inner" /> lies within <paramref name="outer" />.</returns>
    public static bool IsFieldWithinField(Field inner, Field outer, Dictionary<Node, int> order)
    {
        if (ReferenceEquals(inner, outer)) return false;
        if (inner.Start == null || inner.End == null || outer.Start == null || outer.End == null)
            return false;

        if (!order.TryGetValue(inner.Start, out var innerStart)) return false;
        if (!order.TryGetValue(inner.End, out var innerEnd)) return false;
        if (!order.TryGetValue(outer.Start, out var outerStart)) return false;
        if (!order.TryGetValue(outer.End, out var outerEnd)) return false;

        return innerStart > outerStart && innerEnd < outerEnd;
    }

    /// <summary>
    ///     Numbers every node in the order the document reads.
    ///     <para>
    ///         Containment used to be decided by walking siblings from a field's start to its end,
    ///         which only works while both markers share a parent. A field whose result spans a
    ///         paragraph break has them in different paragraphs, so a dangerous field inside it was
    ///         reported as not contained (R7-W01).
    ///     </para>
    /// </summary>
    /// <param name="document">The document to number.</param>
    /// <returns>Each node's position in document order.</returns>
    public static Dictionary<Node, int> DocumentOrder(Document document)
    {
        var order = new Dictionary<Node, int>();
        var position = 0;
        foreach (var node in document.GetChildNodes(NodeType.Any, true))
            order[node] = position++;

        return order;
    }

    /// <summary>
    ///     Whether any inner interval lies strictly inside any outer one, in one sweep.
    /// </summary>
    /// <param name="outers">Intervals that may contain, as (start, end) positions.</param>
    /// <param name="inners">Intervals that must not be contained.</param>
    /// <returns>The first contained inner interval's index, or -1.</returns>
    /// <remarks>
    ///     The nested-field check compared every updating field with every disallowed one, and a
    ///     document decides how many of each it has (R21-RES03). Sorted by start, an inner interval
    ///     is contained by some outer one exactly when an outer that started before it is still
    ///     open past its end — so the sweep keeps the ends of the outers it has entered and asks
    ///     for the largest. Ties on the start position are not containment: strictness is kept.
    /// </remarks>
    public static int FirstContained(IReadOnlyList<(int Start, int End)> outers,
        IReadOnlyList<(int Start, int End)> inners)
    {
        if (outers.Count == 0 || inners.Count == 0) return -1;

        var outerByStart = outers.Select(o => (o.Start, o.End)).OrderBy(o => o.Start).ToList();
        var innerByStart = inners.Select((o, i) => (o.Start, o.End, Index: i)).OrderBy(o => o.Start).ToList();

        // Ends of every outer whose start is behind the sweep. The largest is the only one that
        // matters: if it does not reach past the inner's end, none of them do.
        var openEnds = new PriorityQueue<int, int>();
        var next = 0;

        foreach (var (start, end, index) in innerByStart)
        {
            while (next < outerByStart.Count && outerByStart[next].Start < start)
            {
                openEnds.Enqueue(outerByStart[next].End, -outerByStart[next].End);
                next++;
            }

            if (openEnds.TryPeek(out var farthestEnd, out _) && farthestEnd > end) return index;
        }

        return -1;
    }

    /// <summary>
    ///     Where every field in one document begins and ends, by document position.
    ///     <para>
    ///         Built once and asked many times: the handlers that ask "is this run inside a field"
    ///         ask it per run, and numbering the document for each of those would be quadratic.
    ///     </para>
    /// </summary>
    public sealed class FieldExtents
    {
        private readonly List<(int Start, int End, Field Field)> _fields;
        private readonly Dictionary<Node, int> _order;

        /// <summary>Creates an index over a document's numbered nodes and field extents.</summary>
        /// <param name="order">Positions from <see cref="DocumentOrder" />.</param>
        /// <param name="fields">Each field's start and end position.</param>
        private FieldExtents(Dictionary<Node, int> order, List<(int, int, Field)> fields)
        {
            _order = order;
            _fields = fields;
        }

        /// <summary>Indexes one document's fields.</summary>
        /// <param name="document">The document to index.</param>
        /// <returns>The index.</returns>
        public static FieldExtents Of(Document document)
        {
            var order = DocumentOrder(document);
            var fields = new List<(int, int, Field)>();

            foreach (var field in document.Range.Fields)
            {
                if (field.Start == null || field.End == null) continue;
                if (!order.TryGetValue(field.Start, out var start)) continue;
                if (!order.TryGetValue(field.End, out var end)) continue;

                fields.Add((start, end, field));
            }

            return new FieldExtents(order, fields);
        }

        /// <summary>
        ///     Whether this index numbered the given node.
        ///     <para>
        ///         A caller that keeps an index across edits needs to know when it has gone stale:
        ///         a node the index has never seen gets no answer here, and "no answer" must not
        ///         be read as "not in a field".
        ///     </para>
        /// </summary>
        /// <param name="node">The node to look for.</param>
        /// <returns><c>true</c> when the node was present when the index was built.</returns>
        public bool Knows(Node node)
        {
            return _order.ContainsKey(node);
        }

        /// <summary>
        ///     The innermost field containing a node, or null when it is inside none.
        /// </summary>
        /// <param name="node">The node to test.</param>
        /// <returns>The innermost enclosing field, or null.</returns>
        /// <remarks>
        ///     The range runs from the FieldStart inclusive to the FieldEnd exclusive, which is
        ///     the boundary this helper has always drawn: the end marker closes a field rather
        ///     than living inside it.
        /// </remarks>
        public Field? EnclosingField(Node node)
        {
            if (!_order.TryGetValue(node, out var position)) return null;

            Field? innermost = null;
            var innermostStart = -1;

            foreach (var (start, end, field) in _fields)
                if (start <= position && position < end && start > innermostStart)
                {
                    innermost = field;
                    innermostStart = start;
                }

            return innermost;
        }

        /// <summary>
        ///     Whether removing or replacing this node whole would cut a field in half.
        ///     <para>
        ///         A field that lies entirely inside the node goes with it, which is a complete
        ///         removal and leaves the document valid. A field that starts before the node and
        ///         ends inside it, or starts inside and ends after, is a different matter: taking
        ///         the node removes one of its markers and leaves the other orphaned, so the
        ///         document is left with a field that has no beginning or no end (§21.3).
        ///     </para>
        ///     <para>
        ///         This is what the three text handlers needed and did not have. They removed whole
        ///         paragraphs, or appended a run to one, without asking whether a field ran through
        ///         them — safe for a field inside one paragraph, wrong for one that spans several.
        ///     </para>
        /// </summary>
        /// <param name="node">The node about to be removed or appended to.</param>
        /// <returns><c>true</c> when a field crosses this node's boundary.</returns>
        public bool WouldSplitAField(Node node)
        {
            if (!_order.TryGetValue(node, out var from)) return false;

            var to = from;
            if (node is CompositeNode composite)
                foreach (var descendant in composite.GetChildNodes(NodeType.Any, true))
                    if (_order.TryGetValue(descendant, out var position) && position > to)
                        to = position;

            return _fields.Any(field =>
            {
                var intersects = field.Start <= to && from <= field.End;
                var containedWhole = from <= field.Start && field.End <= to;
                return intersects && !containedWhole;
            });
        }
    }
}
