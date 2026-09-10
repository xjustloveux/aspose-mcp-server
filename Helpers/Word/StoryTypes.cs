namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Canonical document-story names used by paragraph addressing. The story determines which
///     collection a paragraph index is relative to.
/// </summary>
public static class StoryTypes
{
    /// <summary>The main document body, which is the default story.</summary>
    public const string Body = "Body";

    /// <summary>A section's header, addressed together with its header/footer type.</summary>
    public const string Header = "Header";

    /// <summary>A section's footer, addressed together with its header/footer type.</summary>
    public const string Footer = "Footer";

    /// <summary>The paragraphs inside an inline text box, which the body does not contain.</summary>
    public const string TextBox = "TextBox";

    /// <summary>The paragraphs of a comment.</summary>
    public const string Comment = "Comment";

    /// <summary>The paragraphs of a footnote.</summary>
    public const string Footnote = "Footnote";

    /// <summary>The paragraphs of an endnote.</summary>
    public const string Endnote = "Endnote";
}
