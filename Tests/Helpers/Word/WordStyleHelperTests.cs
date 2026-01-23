using Aspose.Words;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers.Word;

public class WordStyleHelperTests : WordTestBase
{
    #region ApplyStyleToParagraph Tests - Invalid Cases

    [Fact]
    public void ApplyStyleToParagraph_WithNullParagraph_ThrowsArgumentNullException()
    {
        var doc = new Document();
        var style = doc.Styles[StyleIdentifier.Normal];

        var ex = Assert.Throws<ArgumentNullException>(() =>
            WordStyleHelper.ApplyStyleToParagraph(null, style, "Normal"));

        Assert.Contains("para", ex.ParamName);
    }

    #endregion

    #region ApplyStyleToParagraph Tests - Valid Cases

    [Fact]
    public void ApplyStyleToParagraph_WithNonEmptyParagraph_AppliesStyle()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test paragraph");
        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().First();
        var style = doc.Styles[StyleIdentifier.Heading1];

        WordStyleHelper.ApplyStyleToParagraph(para, style, "Heading 1");

        Assert.Equal("Heading 1", para.ParagraphFormat.StyleName);
    }

    [Fact]
    public void ApplyStyleToParagraph_WithEmptyParagraph_AppliesStyle()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertParagraph();
        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().First();
        var style = doc.Styles[StyleIdentifier.Heading2];

        WordStyleHelper.ApplyStyleToParagraph(para, style, "Heading 2");

        Assert.Equal("Heading 2", para.ParagraphFormat.StyleName);
    }

    [Fact]
    public void ApplyStyleToParagraph_WithNormalStyle_AppliesStyle()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test");
        var para = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().First();
        var style = doc.Styles[StyleIdentifier.Normal];

        WordStyleHelper.ApplyStyleToParagraph(para, style, "Normal");

        Assert.Equal("Normal", para.ParagraphFormat.StyleName);
    }

    #endregion

    #region CopyStyleProperties Tests

    [Fact]
    public void CopyStyleProperties_CopiesFontProperties()
    {
        var doc = new Document();
        var sourceStyle = doc.Styles.Add(StyleType.Paragraph, "SourceStyle");
        sourceStyle.Font.Name = "Arial";
        sourceStyle.Font.Size = 14;
        sourceStyle.Font.Bold = true;
        sourceStyle.Font.Italic = true;

        var targetStyle = doc.Styles.Add(StyleType.Paragraph, "TargetStyle");

        WordStyleHelper.CopyStyleProperties(sourceStyle, targetStyle);

        Assert.Equal("Arial", targetStyle.Font.Name);
        Assert.Equal(14, targetStyle.Font.Size);
        Assert.True(targetStyle.Font.Bold);
        Assert.True(targetStyle.Font.Italic);
    }

    [Fact]
    public void CopyStyleProperties_ForParagraphStyle_CopiesParagraphFormat()
    {
        var doc = new Document();
        var sourceStyle = doc.Styles.Add(StyleType.Paragraph, "SourcePara");
        sourceStyle.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        sourceStyle.ParagraphFormat.SpaceBefore = 12;
        sourceStyle.ParagraphFormat.SpaceAfter = 6;
        sourceStyle.ParagraphFormat.LeftIndent = 20;

        var targetStyle = doc.Styles.Add(StyleType.Paragraph, "TargetPara");

        WordStyleHelper.CopyStyleProperties(sourceStyle, targetStyle);

        Assert.Equal(ParagraphAlignment.Center, targetStyle.ParagraphFormat.Alignment);
        Assert.Equal(12, targetStyle.ParagraphFormat.SpaceBefore);
        Assert.Equal(6, targetStyle.ParagraphFormat.SpaceAfter);
        Assert.Equal(20, targetStyle.ParagraphFormat.LeftIndent);
    }

    [Fact]
    public void CopyStyleProperties_ForCharacterStyle_OnlyCopiesFontProperties()
    {
        var doc = new Document();
        var sourceStyle = doc.Styles.Add(StyleType.Character, "SourceChar");
        sourceStyle.Font.Name = "Times New Roman";
        sourceStyle.Font.Size = 12;

        var targetStyle = doc.Styles.Add(StyleType.Character, "TargetChar");

        WordStyleHelper.CopyStyleProperties(sourceStyle, targetStyle);

        Assert.Equal("Times New Roman", targetStyle.Font.Name);
        Assert.Equal(12, targetStyle.Font.Size);
    }

    [Fact]
    public void CopyStyleProperties_CopiesUnderline()
    {
        var doc = new Document();
        var sourceStyle = doc.Styles.Add(StyleType.Character, "Source");
        sourceStyle.Font.Underline = Underline.Single;

        var targetStyle = doc.Styles.Add(StyleType.Character, "Target");

        WordStyleHelper.CopyStyleProperties(sourceStyle, targetStyle);

        Assert.Equal(Underline.Single, targetStyle.Font.Underline);
    }

    #endregion
}
