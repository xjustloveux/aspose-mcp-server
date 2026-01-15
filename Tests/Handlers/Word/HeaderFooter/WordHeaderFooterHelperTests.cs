using Aspose.Words;
using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Tests.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class WordHeaderFooterHelperTests : WordTestBase
{
    #region GetHeaderFooterType Tests

    [Theory]
    [InlineData("primary", true, HeaderFooterType.HeaderPrimary)]
    [InlineData("PRIMARY", true, HeaderFooterType.HeaderPrimary)]
    [InlineData("first", true, HeaderFooterType.HeaderFirst)]
    [InlineData("FIRST", true, HeaderFooterType.HeaderFirst)]
    [InlineData("even", true, HeaderFooterType.HeaderEven)]
    [InlineData("EVEN", true, HeaderFooterType.HeaderEven)]
    [InlineData("primary", false, HeaderFooterType.FooterPrimary)]
    [InlineData("PRIMARY", false, HeaderFooterType.FooterPrimary)]
    [InlineData("first", false, HeaderFooterType.FooterFirst)]
    [InlineData("FIRST", false, HeaderFooterType.FooterFirst)]
    [InlineData("even", false, HeaderFooterType.FooterEven)]
    [InlineData("EVEN", false, HeaderFooterType.FooterEven)]
    public void GetHeaderFooterType_WithValidValues_ReturnsCorrectType(string type, bool isHeader,
        HeaderFooterType expected)
    {
        var result = WordHeaderFooterHelper.GetHeaderFooterType(type, isHeader);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid", true, HeaderFooterType.HeaderPrimary)]
    [InlineData("unknown", true, HeaderFooterType.HeaderPrimary)]
    [InlineData("", true, HeaderFooterType.HeaderPrimary)]
    [InlineData("invalid", false, HeaderFooterType.FooterPrimary)]
    [InlineData("unknown", false, HeaderFooterType.FooterPrimary)]
    [InlineData("", false, HeaderFooterType.FooterPrimary)]
    public void GetHeaderFooterType_WithInvalidValues_ReturnsPrimary(string type, bool isHeader,
        HeaderFooterType expected)
    {
        var result = WordHeaderFooterHelper.GetHeaderFooterType(type, isHeader);

        Assert.Equal(expected, result);
    }

    #endregion

    #region GetOrCreateHeaderFooter Tests

    [Fact]
    public void GetOrCreateHeaderFooter_WithNoExisting_CreatesNew()
    {
        var doc = new Document();
        var section = doc.Sections[0];

        var result = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, HeaderFooterType.HeaderPrimary);

        Assert.NotNull(result);
        Assert.Equal(HeaderFooterType.HeaderPrimary, result.HeaderFooterType);
    }

    [SkippableFact]
    public void GetOrCreateHeaderFooter_WithExisting_ReturnsExisting()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode blocks HeaderFooter manipulation");

        var doc = new Document();
        var section = doc.Sections[0];
        var existing = new Aspose.Words.HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        section.HeadersFooters.Add(existing);

        var result = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, HeaderFooterType.HeaderPrimary);

        Assert.Same(existing, result);
    }

    [SkippableFact]
    public void GetOrCreateHeaderFooter_WithDifferentType_CreatesNewForType()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode blocks HeaderFooter manipulation");

        var doc = new Document();
        var section = doc.Sections[0];
        var header = new Aspose.Words.HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        section.HeadersFooters.Add(header);

        var result = WordHeaderFooterHelper.GetOrCreateHeaderFooter(section, doc, HeaderFooterType.FooterPrimary);

        Assert.NotNull(result);
        Assert.Equal(HeaderFooterType.FooterPrimary, result.HeaderFooterType);
    }

    #endregion

    #region ClearTextOnly Tests

    [SkippableFact]
    public void ClearTextOnly_WithTextContent_ClearsText()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode blocks HeaderFooter manipulation");

        var doc = new Document();
        var section = doc.Sections[0];
        var header = new Aspose.Words.HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        section.HeadersFooters.Add(header);
        var para = new WordParagraph(doc);
        header.AppendChild(para);
        var run = new Run(doc, "Header text");
        para.AppendChild(run);

        WordHeaderFooterHelper.ClearTextOnly(header);

        Assert.Equal(string.Empty, run.Text);
    }

    [SkippableFact]
    public void ClearTextOnly_WithEmptyHeaderFooter_DoesNotThrow()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode blocks HeaderFooter manipulation");

        var doc = new Document();
        var section = doc.Sections[0];
        var header = new Aspose.Words.HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        section.HeadersFooters.Add(header);

        var exception = Record.Exception(() => WordHeaderFooterHelper.ClearTextOnly(header));

        Assert.Null(exception);
    }

    #endregion

    #region InsertTextOrField Tests

    [Fact]
    public void InsertTextOrField_WithPlainText_InsertsText()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        WordHeaderFooterHelper.InsertTextOrField(builder, "Hello World", null, null, null, null);

        Assert.Contains("Hello World", doc.GetText());
    }

    [Fact]
    public void InsertTextOrField_WithPageField_InsertsField()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        WordHeaderFooterHelper.InsertTextOrField(builder, "{PAGE}", null, null, null, null);

        var fields = doc.Range.Fields;
        Assert.True(fields.Count > 0);
    }

    [Fact]
    public void InsertTextOrField_WithFontSettings_AppliesFont()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        WordHeaderFooterHelper.InsertTextOrField(builder, "Styled Text", "Arial", null, null, 14);

        Assert.Equal("Arial", builder.Font.Name);
        Assert.Equal(14, builder.Font.Size);
    }

    [Theory]
    [InlineData("{NUMPAGES}")]
    [InlineData("{DATE}")]
    [InlineData("{TIME}")]
    [InlineData("{FILENAME}")]
    [InlineData("{AUTHOR}")]
    [InlineData("{TITLE}")]
    public void InsertTextOrField_WithKnownFieldCodes_InsertsField(string fieldCode)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        WordHeaderFooterHelper.InsertTextOrField(builder, fieldCode, null, null, null, null);

        var fields = doc.Range.Fields;
        Assert.True(fields.Count > 0);
    }

    [Fact]
    public void InsertTextOrField_WithCustomFieldCode_InsertsCustomField()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        WordHeaderFooterHelper.InsertTextOrField(builder, "{CUSTOMFIELD}", null, null, null, null);

        var fields = doc.Range.Fields;
        Assert.True(fields.Count > 0);
    }

    #endregion
}
