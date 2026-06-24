using System.Runtime.CompilerServices;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     The single entry point for resolving a <see cref="ParagraphAddress" /> to a paragraph node.
///     All Word handlers locate paragraphs through this resolver so the paragraph-index space stays
///     consistent across every tool.
/// </summary>
public static class ParagraphResolver
{
    /// <summary>
    ///     Per-document registry of stable paragraph handles. Keyed by document instance so the
    ///     handles live exactly as long as the in-memory document and are garbage-collected with it.
    /// </summary>
    private static readonly ConditionalWeakTable<Document, HandleRegistry> HandleTables = new();

    private const string Primary = "Primary";
    private const string First = "First";
    private const string Even = "Even";

    /// <summary>
    ///     Resolves an address to a live paragraph node within the document.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="address">The paragraph address.</param>
    /// <returns>The resolved paragraph, its normalized address, and its document-order index.</returns>
    /// <exception cref="ArgumentException">Thrown when the address is out of range or names a missing story.</exception>
    public static ParagraphRef Resolve(Document doc, ParagraphAddress address)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(address);

        if (!string.IsNullOrEmpty(address.Handle))
            return ResolveByHandle(doc, address.Handle);

        var paragraphs = GetStoryParagraphs(doc, address);
        var resolvedIndex = address.Index == -1 ? paragraphs.Count - 1 : address.Index;
        if (resolvedIndex < 0 || resolvedIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"paragraphIndex {address.Index} is out of range for story '{address.StoryType}' " +
                $"(it has {paragraphs.Count} paragraphs).");

        var para = paragraphs[resolvedIndex];
        var documentOrderIndex = doc.GetChildNodes(NodeType.Paragraph, true).IndexOf(para);
        return new ParagraphRef(para, address with { Index = resolvedIndex }, documentOrderIndex);
    }

    /// <summary>
    ///     Computes the address of an existing paragraph node (the reverse of <see cref="Resolve" />),
    ///     so result-emitting operations can report a story-relative address callers can reuse.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="paragraph">The paragraph to address.</param>
    /// <returns>The paragraph, its address, and its document-order index.</returns>
    /// <exception cref="ArgumentException">Thrown when the paragraph does not belong to the document.</exception>
    public static ParagraphRef AddressOf(Document doc, WordParagraph paragraph)
    {
        ArgumentNullException.ThrowIfNull(doc);
        return AddressOf(doc, paragraph, new AddressingContext(doc));
    }

    /// <summary>
    ///     Computes the address of a paragraph reusing a caller-built <see cref="AddressingContext" /> so
    ///     that a bulk operation addressing many paragraphs of one document builds each paragraph-index
    ///     map once instead of rescanning per paragraph (turns an O(n²) loop into O(n)).
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="paragraph">The paragraph to address.</param>
    /// <param name="context">The addressing context memoizing this document's paragraph-index maps.</param>
    /// <returns>The paragraph, its address, and its document-order index.</returns>
    /// <exception cref="ArgumentException">Thrown when the paragraph does not belong to the document.</exception>
    public static ParagraphRef AddressOf(Document doc, WordParagraph paragraph, AddressingContext context)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(paragraph);
        ArgumentNullException.ThrowIfNull(context);
        if (!ReferenceEquals(doc, context.Document))
            throw new ArgumentException(
                "The addressing context was built for a different document.", nameof(context));

        var (address, storyContainer) = ClassifyStory(doc, paragraph, context);
        var index = context.StoryIndexOf(storyContainer, paragraph);
        if (index < 0)
            throw new ArgumentException(
                "The paragraph does not belong to this document (it is not addressable within its story).",
                nameof(paragraph));
        var documentOrderIndex = context.DocumentOrderIndexOf(paragraph);
        return new ParagraphRef(paragraph, address with { Index = index }, documentOrderIndex);
    }

    /// <summary>
    ///     Mints (or returns the existing) stable handle for a paragraph within the given document
    ///     instance. A handle survives index shifts and stays valid for the life of the in-memory
    ///     document — i.e. across calls in a session, where the same <see cref="Document" /> instance
    ///     is reused. Handles are keyed by document instance, so they do not carry across a file-mode
    ///     reload (which produces a different instance).
    /// </summary>
    /// <param name="doc">The document the paragraph belongs to.</param>
    /// <param name="paragraph">The paragraph to mint a handle for.</param>
    /// <returns>The opaque handle string.</returns>
    public static string MintHandle(Document doc, WordParagraph paragraph)
    {
        ArgumentNullException.ThrowIfNull(doc);
        ArgumentNullException.ThrowIfNull(paragraph);
        return HandleTables.GetOrCreateValue(doc).Mint(paragraph);
    }

    /// <summary>
    ///     Resolves a handle to its live paragraph, reporting the paragraph's current address (so a
    ///     handle minted before an edit still resolves after indices shift). Throws when the handle is
    ///     unknown to this document instance or the paragraph it named has since been removed.
    /// </summary>
    private static ParagraphRef ResolveByHandle(Document doc, string handle)
    {
        var node = HandleTables.TryGetValue(doc, out var registry) ? registry.Lookup(handle) : null;
        if (node == null)
            throw new ArgumentException(
                $"Unknown paragraph handle '{handle}'. Handles are only valid within the session that " +
                "issued them; re-fetch the paragraph to obtain a current handle.");
        if (node is not WordParagraph para || para.ParentNode == null)
            throw new ArgumentException(
                $"Paragraph handle '{handle}' is stale (the paragraph it named was removed). Re-fetch the " +
                "paragraph to obtain a current handle.");
        return AddressOf(doc, para);
    }

    /// <summary>
    ///     Classifies which story a paragraph belongs to and returns the story address (with
    ///     container-relative index left at 0, filled by the caller) plus the container node to
    ///     index the paragraph within.
    /// </summary>
    /// <param name="doc">The document the paragraph belongs to.</param>
    /// <param name="paragraph">The paragraph to classify.</param>
    /// <param name="context">The addressing context providing memoized container and index lookups.</param>
    /// <returns>The story address (index left at 0) and the container node to index the paragraph within.</returns>
    private static (ParagraphAddress Address, CompositeNode Container) ClassifyStory(Document doc,
        WordParagraph paragraph, AddressingContext context)
    {
        if (paragraph.GetAncestor(NodeType.Comment) is Comment comment)
            return (new ParagraphAddress(0, StoryTypes.Comment, ContainerIndex: comment.Id), comment);

        if (paragraph.GetAncestor(NodeType.Shape) is Shape shape)
        {
            var shapeIndex = context.TextBoxShapes().IndexOf(shape);
            if (shapeIndex < 0)
                throw new ArgumentException(
                    "The paragraph's text-box shape does not belong to this document.", nameof(paragraph));
            return (new ParagraphAddress(0, StoryTypes.TextBox, ContainerIndex: shapeIndex), shape);
        }

        if (paragraph.GetAncestor(NodeType.Footnote) is Footnote footnote)
        {
            var isEndnote = footnote.FootnoteType == FootnoteType.Endnote;
            var noteStory = isEndnote ? StoryTypes.Endnote : StoryTypes.Footnote;
            var noteIndex = context.Notes(isEndnote).IndexOf(footnote);
            if (noteIndex < 0)
                throw new ArgumentException(
                    "The paragraph's footnote/endnote does not belong to this document.", nameof(paragraph));
            return (new ParagraphAddress(0, noteStory, ContainerIndex: noteIndex), footnote);
        }

        if (paragraph.GetAncestor(NodeType.HeaderFooter) is HeaderFooter headerFooter)
        {
            var (storyType, hfType) = MapHeaderFooter(headerFooter.HeaderFooterType);
            return (new ParagraphAddress(0, storyType, SectionIndexOf(doc, headerFooter), hfType), headerFooter);
        }

        var section = paragraph.GetAncestor(NodeType.Section) as Section ?? doc.FirstSection;
        return (new ParagraphAddress(0, StoryTypes.Body, SectionIndexOf(doc, section)), section.Body);
    }

    /// <summary>
    ///     Maps an Aspose <see cref="HeaderFooterType" /> to the canonical story name and
    ///     Primary/First/Even header-footer type used by paragraph addressing.
    /// </summary>
    /// <param name="type">The Aspose header/footer type.</param>
    /// <returns>The story type (Header or Footer) and the Primary/First/Even discriminator.</returns>
    private static (string StoryType, string HeaderFooterType) MapHeaderFooter(HeaderFooterType type)
    {
        return type switch
        {
            HeaderFooterType.HeaderPrimary => (StoryTypes.Header, Primary),
            HeaderFooterType.HeaderFirst => (StoryTypes.Header, First),
            HeaderFooterType.HeaderEven => (StoryTypes.Header, Even),
            HeaderFooterType.FooterPrimary => (StoryTypes.Footer, Primary),
            HeaderFooterType.FooterFirst => (StoryTypes.Footer, First),
            HeaderFooterType.FooterEven => (StoryTypes.Footer, Even),
            _ => (StoryTypes.Header, Primary)
        };
    }

    /// <summary>
    ///     Returns the index of the section that contains the given node (0 when it cannot be determined).
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="node">The node whose section index is wanted (or the section itself).</param>
    /// <returns>The 0-based section index, or 0 when no containing section is found.</returns>
    private static int SectionIndexOf(Document doc, Node? node)
    {
        var section = node as Section ?? node?.GetAncestor(NodeType.Section) as Section;
        return section == null ? 0 : doc.Sections.IndexOf(section);
    }

    /// <summary>
    ///     Returns the paragraphs of the story named by the address (Body of a section, or a
    ///     Header/Footer), in story order. Range operations that need the whole story list resolve
    ///     through this so they share the resolver's index space. The address's Index is ignored.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="address">The address whose story selects the paragraph collection.</param>
    /// <returns>The story's paragraphs.</returns>
    public static List<WordParagraph> GetStoryParagraphs(Document doc, ParagraphAddress address)
    {
        switch (address.StoryType)
        {
            case StoryTypes.Body:
                return GetSection(doc, address.SectionIndex).Body
                    .GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            case StoryTypes.Header:
            case StoryTypes.Footer:
                var headerFooter = GetSection(doc, address.SectionIndex)
                    .HeadersFooters[ResolveHeaderFooterType(address.StoryType, address.HeaderFooterType)];
                if (headerFooter == null)
                    throw new ArgumentException(
                        $"Section {address.SectionIndex} has no {address.StoryType} '{address.HeaderFooterType}'.");
                return headerFooter.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            case StoryTypes.TextBox:
                return GetContainerParagraphs(GetTextBoxShapes(doc).Cast<CompositeNode>().ToList(),
                    address.ContainerIndex ?? 0, "text boxes");
            case StoryTypes.Comment:
                return GetCommentParagraphs(doc, address);
            case StoryTypes.Footnote:
            case StoryTypes.Endnote:
                return GetNoteParagraphs(doc, address);
            default:
                throw new NotSupportedException(
                    $"storyType '{address.StoryType}' is not supported (supported: Body, Header, Footer, " +
                    "TextBox, Comment, Footnote, Endnote).");
        }
    }

    /// <summary>
    ///     Returns every text-bearing shape (text box) in document order, the container set that
    ///     TextBox addresses index into.
    /// </summary>
    private static List<Shape> GetTextBoxShapes(Document doc)
    {
        return doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.GetChildNodes(NodeType.Paragraph, true).Count > 0)
            .ToList();
    }

    /// <summary>
    ///     Returns the footnotes or endnotes in document order (the set a note container index selects).
    /// </summary>
    private static List<Footnote> GetNotes(Document doc, bool endnotes)
    {
        return doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote == endnotes)
            .ToList();
    }

    /// <summary>
    ///     Selects a container by ordinal index and returns its paragraphs.
    /// </summary>
    private static List<WordParagraph> GetContainerParagraphs(IReadOnlyList<CompositeNode> containers,
        int containerIndex, string label)
    {
        if (containers.Count == 0)
            throw new ArgumentException($"Document has no {label}.");
        if (containerIndex < 0 || containerIndex >= containers.Count)
            throw new ArgumentException(
                $"containerIndex {containerIndex} is out of range (document has {containers.Count} {label}).");
        return containers[containerIndex].GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
    }

    /// <summary>
    ///     Returns the paragraphs of the comment selected by the address. For Comment stories the
    ///     container index is the comment's stable id, not a positional ordinal.
    /// </summary>
    private static List<WordParagraph> GetCommentParagraphs(Document doc, ParagraphAddress address)
    {
        var comments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        if (comments.Count == 0)
            throw new ArgumentException("Document has no comments.");

        var comment = address.ContainerIndex.HasValue
            ? comments.FirstOrDefault(c => c.Id == address.ContainerIndex.Value)
              ?? throw new ArgumentException(
                  $"No comment has id {address.ContainerIndex.Value} " +
                  "(for Comment stories, containerIndex is the comment id).")
            : comments[0];

        return comment.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
    }

    /// <summary>
    ///     Returns the paragraphs of the footnote / endnote selected by the address.
    /// </summary>
    private static List<WordParagraph> GetNoteParagraphs(Document doc, ParagraphAddress address)
    {
        var endnotes = address.StoryType == StoryTypes.Endnote;
        var notes = GetNotes(doc, endnotes).Cast<CompositeNode>().ToList();
        return GetContainerParagraphs(notes, address.ContainerIndex ?? 0, endnotes ? "endnotes" : "footnotes");
    }

    private static Section GetSection(Document doc, int sectionIndex)
    {
        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
            throw new ArgumentException(
                $"sectionIndex {sectionIndex} is out of range (document has {doc.Sections.Count} sections).");
        return doc.Sections[sectionIndex];
    }

    private static HeaderFooterType ResolveHeaderFooterType(string storyType, string headerFooterType)
    {
        var isHeader = storyType == StoryTypes.Header;
        return headerFooterType switch
        {
            Primary => isHeader ? HeaderFooterType.HeaderPrimary : HeaderFooterType.FooterPrimary,
            First => isHeader ? HeaderFooterType.HeaderFirst : HeaderFooterType.FooterFirst,
            Even => isHeader ? HeaderFooterType.HeaderEven : HeaderFooterType.FooterEven,
            _ => throw new ArgumentException(
                $"headerFooterType '{headerFooterType}' is invalid (expected Primary, First, or Even).")
        };
    }

    /// <summary>
    ///     Per-operation memoization of a document's paragraph-index lookups. A bulk operation that calls
    ///     the context-taking <c>AddressOf</c> overload for many paragraphs builds each document-order /
    ///     story-container index map once here instead of per paragraph. Build one per operation and
    ///     discard it; it assumes the document is not structurally mutated while in use.
    /// </summary>
    public sealed class AddressingContext
    {
        private readonly Dictionary<CompositeNode, Dictionary<Node, int>> _containerOrder =
            new(ReferenceEqualityComparer.Instance);

        private Dictionary<Node, int>? _documentOrder;
        private List<Footnote>? _endnotes;
        private List<Footnote>? _footnotes;
        private List<Shape>? _textBoxShapes;

        /// <summary>
        ///     Creates a context bound to a document.
        /// </summary>
        /// <param name="doc">The document the addressed paragraphs belong to.</param>
        public AddressingContext(Document doc)
        {
            ArgumentNullException.ThrowIfNull(doc);
            Document = doc;
        }

        /// <summary>
        ///     The document this context is bound to. Paragraphs addressed with this context must belong
        ///     to this same document instance.
        /// </summary>
        internal Document Document { get; }

        internal int DocumentOrderIndexOf(Node paragraph)
        {
            _documentOrder ??= BuildIndexMap(Document.GetChildNodes(NodeType.Paragraph, true));
            return _documentOrder.GetValueOrDefault(paragraph, -1);
        }

        internal int StoryIndexOf(CompositeNode container, Node paragraph)
        {
            if (!_containerOrder.TryGetValue(container, out var map))
                _containerOrder[container] = map = BuildIndexMap(container.GetChildNodes(NodeType.Paragraph, true));
            return map.GetValueOrDefault(paragraph, -1);
        }

        internal List<Shape> TextBoxShapes()
        {
            return _textBoxShapes ??= GetTextBoxShapes(Document);
        }

        internal List<Footnote> Notes(bool endnotes)
        {
            return endnotes ? _endnotes ??= GetNotes(Document, true) : _footnotes ??= GetNotes(Document, false);
        }

        private static Dictionary<Node, int> BuildIndexMap(NodeCollection nodes)
        {
            var map = new Dictionary<Node, int>(ReferenceEqualityComparer.Instance);
            var i = 0;
            foreach (var node in nodes) map[node] = i++;
            return map;
        }
    }

    private sealed class HandleRegistry
    {
        private readonly Dictionary<string, Node> _byHandle = new();
        private readonly Dictionary<Node, string> _byNode = new(ReferenceEqualityComparer.Instance);
        private readonly object _gate = new();
        private int _counter;

        public string Mint(Node node)
        {
            lock (_gate)
            {
                if (_byNode.TryGetValue(node, out var existing))
                    return existing;
                var handle = "p" + _counter++;
                _byHandle[handle] = node;
                _byNode[node] = handle;
                return handle;
            }
        }

        public Node? Lookup(string handle)
        {
            lock (_gate)
            {
                return _byHandle.GetValueOrDefault(handle);
            }
        }
    }
}
