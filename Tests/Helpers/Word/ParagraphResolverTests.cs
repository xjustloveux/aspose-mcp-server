using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Helpers.Word;

public class ParagraphResolverTests
{
    private static Document TwoBodyParagraphs()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Writeln("body0");
        b.Write("body1");
        return doc;
    }

    [Fact]
    public void Resolve_Body_ReturnsParagraphByIndex()
    {
        var doc = TwoBodyParagraphs();
        var r = ParagraphResolver.Resolve(doc, new ParagraphAddress(0));
        Assert.StartsWith("body0", r.Paragraph.GetText());
        Assert.Equal(0, r.Address.Index);
        Assert.Equal(StoryTypes.Body, r.Address.StoryType);
    }

    [Fact]
    public void Resolve_Body_MinusOne_ReturnsLastAndNormalizes()
    {
        var doc = TwoBodyParagraphs();
        var r = ParagraphResolver.Resolve(doc, new ParagraphAddress(-1));
        Assert.StartsWith("body1", r.Paragraph.GetText());
        Assert.Equal(1, r.Address.Index);
    }

    [Fact]
    public void Resolve_Body_OutOfRange_Throws()
    {
        var doc = TwoBodyParagraphs();
        Assert.Throws<ArgumentException>(() => ParagraphResolver.Resolve(doc, new ParagraphAddress(99)));
    }

    [Fact]
    public void Resolve_Header_ReturnsHeaderParagraph()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Write("body");
        b.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        b.Write("hdr");

        var r = ParagraphResolver.Resolve(doc, new ParagraphAddress(0, StoryTypes.Header));
        Assert.StartsWith("hdr", r.Paragraph.GetText());
    }

    [Fact]
    public void Resolve_Footer_ReturnsFooterParagraph()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Write("body");
        b.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        b.Write("ftr");

        var r = ParagraphResolver.Resolve(doc, new ParagraphAddress(0, StoryTypes.Footer));
        Assert.StartsWith("ftr", r.Paragraph.GetText());
    }

    [Fact]
    public void Resolve_Body_ReportsDocumentOrderIndex()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Write("body0");
        b.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        b.Write("hdr");

        var r = ParagraphResolver.Resolve(doc, new ParagraphAddress(0));
        var expected = doc.GetChildNodes(NodeType.Paragraph, true).IndexOf(r.Paragraph);
        Assert.Equal(expected, r.DocumentOrderIndex);
    }

    [Fact]
    public void AddressOf_BodyParagraph_ReturnsBodyAddress()
    {
        var doc = TwoBodyParagraphs();
        var para = doc.FirstSection.Body.Paragraphs[1];
        var r = ParagraphResolver.AddressOf(doc, para);
        Assert.Equal(StoryTypes.Body, r.Address.StoryType);
        Assert.Equal(1, r.Address.Index);
        Assert.Equal(0, r.Address.SectionIndex);
    }

    [Fact]
    public void AddressOf_HeaderParagraph_ReturnsHeaderAddress()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Write("body");
        b.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        b.Write("hdr");
        var headerPara = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].Paragraphs[0];

        var r = ParagraphResolver.AddressOf(doc, headerPara);
        Assert.Equal(StoryTypes.Header, r.Address.StoryType);
        Assert.Equal("Primary", r.Address.HeaderFooterType);
        Assert.Equal(0, r.Address.Index);
    }

    [Fact]
    public void AddressOf_ThenResolve_RoundTripsToSameNode()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Writeln("body0");
        b.Write("body1");
        var para = doc.FirstSection.Body.Paragraphs[1];

        var addr = ParagraphResolver.AddressOf(doc, para).Address;
        var back = ParagraphResolver.Resolve(doc, addr).Paragraph;
        Assert.Same(para, back);
    }

    [Fact]
    public void Resolve_TextBox_RoundTripsToSameNode()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Write("body");
        var shape = b.InsertShape(ShapeType.TextBox, 100, 50);
        var tbPara = new WordParagraph(doc);
        tbPara.AppendChild(new Run(doc, "tb-text"));
        shape.AppendChild(tbPara);

        var addr = ParagraphResolver.AddressOf(doc, tbPara).Address;
        Assert.Equal(StoryTypes.TextBox, addr.StoryType);

        var back = ParagraphResolver.Resolve(doc, addr).Paragraph;
        Assert.Same(tbPara, back);
    }

    [Fact]
    public void Resolve_Comment_RoundTripsAndContainerIndexIsCommentId()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Write("body");
        var comment = new Comment(doc, "Author", "A", DateTime.Now);
        var cPara = new WordParagraph(doc);
        cPara.AppendChild(new Run(doc, "comment-text"));
        comment.AppendChild(cPara);
        doc.FirstSection.Body.FirstParagraph.AppendChild(comment);

        var addr = ParagraphResolver.AddressOf(doc, cPara).Address;
        Assert.Equal(StoryTypes.Comment, addr.StoryType);
        Assert.Equal(comment.Id, addr.ContainerIndex);

        var back = ParagraphResolver.Resolve(doc, addr).Paragraph;
        Assert.Same(cPara, back);
    }

    [Fact]
    public void Resolve_Footnote_RoundTripsToSameNode()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Write("body");
        var footnote = b.InsertFootnote(FootnoteType.Footnote, "footnote-text");
        var fPara = (WordParagraph)footnote.GetChildNodes(NodeType.Paragraph, true)[0];

        var addr = ParagraphResolver.AddressOf(doc, fPara).Address;
        Assert.Equal(StoryTypes.Footnote, addr.StoryType);

        var back = ParagraphResolver.Resolve(doc, addr).Paragraph;
        Assert.Same(fPara, back);
    }

    [Fact]
    public void Resolve_Endnote_RoundTripsToSameNode()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Write("body");
        var endnote = b.InsertFootnote(FootnoteType.Endnote, "endnote-text");
        var ePara = (WordParagraph)endnote.GetChildNodes(NodeType.Paragraph, true)[0];

        var addr = ParagraphResolver.AddressOf(doc, ePara).Address;
        Assert.Equal(StoryTypes.Endnote, addr.StoryType);

        var back = ParagraphResolver.Resolve(doc, addr).Paragraph;
        Assert.Same(ePara, back);
    }

    /// <summary>
    ///     Body: [body0, body1(with inline text box of tb0/tb1), body2, table(cell0), body3].
    ///     The Body story must be body0..body3 + cell0 (tables belong to the body story) and must
    ///     NOT count tb0/tb1 (they belong to the TextBox story).
    /// </summary>
    private static Document BodyWithTextBoxAndTable()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Writeln("body0");
        b.Write("body1");
        var shape = b.InsertShape(ShapeType.TextBox, 100, 50);
        var tb0 = new WordParagraph(doc);
        tb0.AppendChild(new Run(doc, "tb0"));
        shape.AppendChild(tb0);
        var tb1 = new WordParagraph(doc);
        tb1.AppendChild(new Run(doc, "tb1"));
        shape.AppendChild(tb1);
        b.Writeln();
        b.Writeln("body2");
        b.StartTable();
        b.InsertCell();
        b.Write("cell0");
        b.EndRow();
        b.EndTable();
        b.Write("body3");
        return doc;
    }

    [Fact]
    public void GetStoryParagraphs_Body_ExcludesNestedStoryParagraphs_IncludesTableCells()
    {
        var doc = BodyWithTextBoxAndTable();

        var list = ParagraphResolver.GetStoryParagraphs(doc, new ParagraphAddress(0));

        Assert.Equal(5, list.Count);
        Assert.StartsWith("body0", list[0].GetText());
        Assert.StartsWith("body1", list[1].GetText());
        Assert.StartsWith("body2", list[2].GetText());
        Assert.StartsWith("cell0", list[3].GetText());
        Assert.StartsWith("body3", list[4].GetText());
    }

    [Fact]
    public void Resolve_Body_AfterInlineTextBox_ReturnsBodyParagraphNotTextBoxContent()
    {
        var doc = BodyWithTextBoxAndTable();

        var r = ParagraphResolver.Resolve(doc, new ParagraphAddress(2));

        Assert.StartsWith("body2", r.Paragraph.GetText());
    }

    [Fact]
    public void AddressOf_BodyParagraphAfterTextBox_IndexSkipsTextBoxParagraphs()
    {
        var doc = BodyWithTextBoxAndTable();
        var body2 = doc.FirstSection.Body.Paragraphs[2];

        var r = ParagraphResolver.AddressOf(doc, body2);

        Assert.Equal(StoryTypes.Body, r.Address.StoryType);
        Assert.Equal(2, r.Address.Index);
        Assert.Same(body2, ParagraphResolver.Resolve(doc, r.Address).Paragraph);
    }

    [Fact]
    public void AddressOf_TableCellParagraph_IsBodyStoryAndRoundTrips()
    {
        var doc = BodyWithTextBoxAndTable();
        var cell0 = doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.FirstParagraph;

        var r = ParagraphResolver.AddressOf(doc, cell0);

        Assert.Equal(StoryTypes.Body, r.Address.StoryType);
        Assert.Equal(3, r.Address.Index);
        Assert.Same(cell0, ParagraphResolver.Resolve(doc, r.Address).Paragraph);
    }

    [Fact]
    public void TextBoxStory_InBodyWithTable_StillRoundTrips()
    {
        var doc = BodyWithTextBoxAndTable();
        var shape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
        var tb1 = (WordParagraph)shape.GetChildNodes(NodeType.Paragraph, true)[1];

        var r = ParagraphResolver.AddressOf(doc, tb1);

        Assert.Equal(StoryTypes.TextBox, r.Address.StoryType);
        Assert.Equal(1, r.Address.Index);
        Assert.Same(tb1, ParagraphResolver.Resolve(doc, r.Address).Paragraph);
    }

    [Fact]
    public void Handle_MintThenResolve_ReturnsSameNode()
    {
        var doc = TwoBodyParagraphs();
        var para = doc.FirstSection.Body.Paragraphs[1];

        var handle = ParagraphResolver.MintHandle(doc, para);
        var back = ParagraphResolver.Resolve(doc, new ParagraphAddress(0, Handle: handle)).Paragraph;

        Assert.Same(para, back);
    }

    [Fact]
    public void Handle_SurvivesIndexShift_AndReportsCurrentAddress()
    {
        var doc = TwoBodyParagraphs();
        var para = doc.FirstSection.Body.Paragraphs[1];
        var handle = ParagraphResolver.MintHandle(doc, para);

        var inserted = new WordParagraph(doc);
        inserted.AppendChild(new Run(doc, "inserted"));
        doc.FirstSection.Body.InsertBefore(inserted, doc.FirstSection.Body.FirstParagraph);

        var resolved = ParagraphResolver.Resolve(doc, new ParagraphAddress(0, Handle: handle));
        Assert.Same(para, resolved.Paragraph);
        Assert.Equal(2, resolved.Address.Index);
    }

    [Fact]
    public void Handle_Stale_AfterRemoval_Throws()
    {
        var doc = TwoBodyParagraphs();
        var para = doc.FirstSection.Body.Paragraphs[1];
        var handle = ParagraphResolver.MintHandle(doc, para);

        para.Remove();

        Assert.Throws<ArgumentException>(() =>
            ParagraphResolver.Resolve(doc, new ParagraphAddress(0, Handle: handle)));
    }

    [Fact]
    public void Handle_Unknown_Throws()
    {
        var doc = TwoBodyParagraphs();

        Assert.Throws<ArgumentException>(() =>
            ParagraphResolver.Resolve(doc, new ParagraphAddress(0, Handle: "no-such-handle")));
    }

    [Fact]
    public void Handle_MintIsStableForSameNode()
    {
        var doc = TwoBodyParagraphs();
        var para = doc.FirstSection.Body.Paragraphs[0];

        var first = ParagraphResolver.MintHandle(doc, para);
        var second = ParagraphResolver.MintHandle(doc, para);

        Assert.Equal(first, second);
    }

    [Fact]
    public void From_DefaultsToBodySectionZero()
    {
        var p = new OperationParameters();
        var addr = ParagraphAddress.From(p, 5);
        Assert.Equal(5, addr.Index);
        Assert.Equal(StoryTypes.Body, addr.StoryType);
        Assert.Equal(0, addr.SectionIndex);
        Assert.Equal("Primary", addr.HeaderFooterType);
    }

    [Fact]
    public void From_ReadsExplicitStoryFields()
    {
        var p = new OperationParameters();
        p.Set("storyType", StoryTypes.Header);
        p.Set("sectionIndex", 2);
        p.Set("headerFooterType", "First");
        p.Set("containerIndex", 4);

        var addr = ParagraphAddress.From(p, 3);
        Assert.Equal(StoryTypes.Header, addr.StoryType);
        Assert.Equal(2, addr.SectionIndex);
        Assert.Equal("First", addr.HeaderFooterType);
        Assert.Equal(4, addr.ContainerIndex);
    }

    [Fact]
    public void AddressOf_ParagraphNotInDocument_Throws()
    {
        var doc = TwoBodyParagraphs();
        var orphan = new WordParagraph(doc);
        orphan.AppendChild(new Run(doc, "orphan"));

        Assert.Throws<ArgumentException>(() => ParagraphResolver.AddressOf(doc, orphan));
    }

    [Fact]
    public void AddressOf_ParagraphInShapeNotInDocument_Throws()
    {
        var doc = new Document();
        var shape = new Shape(doc, ShapeType.TextBox);
        var tbPara = new WordParagraph(doc);
        tbPara.AppendChild(new Run(doc, "tb"));
        shape.AppendChild(tbPara);

        Assert.Throws<ArgumentException>(() => ParagraphResolver.AddressOf(doc, tbPara));
    }

    [Fact]
    public void Resolve_Body_SectionOutOfRange_Throws()
    {
        var doc = TwoBodyParagraphs();
        Assert.Throws<ArgumentException>(() =>
            ParagraphResolver.Resolve(doc, new ParagraphAddress(0, StoryTypes.Body, 99)));
    }

    [Fact]
    public void AddressOf_HeaderInSecondSection_RoundTripsWithSectionIndex()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Writeln("s0-body");
        b.InsertBreak(BreakType.SectionBreakNewPage);
        b.Writeln("s1-body");
        b.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        b.Write("s1-hdr");
        var headerPara = doc.Sections[1].HeadersFooters[HeaderFooterType.HeaderPrimary].Paragraphs[0];

        var addr = ParagraphResolver.AddressOf(doc, headerPara).Address;
        Assert.Equal(StoryTypes.Header, addr.StoryType);
        Assert.Equal(1, addr.SectionIndex);

        var back = ParagraphResolver.Resolve(doc, addr).Paragraph;
        Assert.Same(headerPara, back);
    }

    [Fact]
    public void AddressOf_SharedContext_MatchesPerCallResults()
    {
        var doc = new Document();
        var b = new DocumentBuilder(doc);
        b.Writeln("body0");
        b.Write("body1");
        b.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        b.Write("hdr");
        b.MoveToDocumentEnd();
        var shape = b.InsertShape(ShapeType.TextBox, 100, 50);
        var tb0 = new WordParagraph(doc);
        tb0.AppendChild(new Run(doc, "tb0"));
        shape.AppendChild(tb0);
        var tb1 = new WordParagraph(doc);
        tb1.AppendChild(new Run(doc, "tb1"));
        shape.AppendChild(tb1);
        b.MoveToDocumentEnd();
        b.InsertFootnote(FootnoteType.Footnote, "fn");

        var paras = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        var ctx = new ParagraphResolver.AddressingContext(doc);

        foreach (var p in paras)
        {
            var perCall = ParagraphResolver.AddressOf(doc, p);
            var shared = ParagraphResolver.AddressOf(doc, p, ctx);
            Assert.Equal(perCall.Address.StoryType, shared.Address.StoryType);
            Assert.Equal(perCall.Address.Index, shared.Address.Index);
            Assert.Equal(perCall.Address.SectionIndex, shared.Address.SectionIndex);
            Assert.Equal(perCall.Address.HeaderFooterType, shared.Address.HeaderFooterType);
            Assert.Equal(perCall.Address.ContainerIndex, shared.Address.ContainerIndex);
            Assert.Equal(perCall.DocumentOrderIndex, shared.DocumentOrderIndex);
        }
    }

    [Fact]
    public void AddressOf_ContextFromDifferentDocument_Throws()
    {
        var doc1 = TwoBodyParagraphs();
        var doc2 = TwoBodyParagraphs();
        var ctxForDoc2 = new ParagraphResolver.AddressingContext(doc2);
        var paraInDoc1 = doc1.FirstSection.Body.Paragraphs[0];

        Assert.Throws<ArgumentException>(() => ParagraphResolver.AddressOf(doc1, paraInDoc1, ctxForDoc2));
    }
}
