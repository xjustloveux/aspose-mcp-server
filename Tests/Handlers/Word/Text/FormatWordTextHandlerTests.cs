using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

public class FormatWordTextHandlerTests : WordHandlerTestBase
{
    private readonly FormatWordTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Format()
    {
        Assert.Equal("format", _handler.Operation);
    }

    #endregion

    #region Color Formatting

    [Theory]
    [InlineData("#FF0000", 255, 0, 0)]
    [InlineData("#00FF00", 0, 255, 0)]
    [InlineData("#0000FF", 0, 0, 255)]
    public void Execute_WithColor_AppliesColor(string color, int r, int g, int b)
    {
        var doc = CreateDocumentWithText("Test text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "color", color }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formatting", result, StringComparison.OrdinalIgnoreCase);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.Equal(r, runs[0].Font.Color.R);
        Assert.Equal(g, runs[0].Font.Color.G);
        Assert.Equal(b, runs[0].Font.Color.B);
        AssertModified(context);
    }

    #endregion

    #region Underline Formatting

    [Theory]
    [InlineData("single", Underline.Single)]
    [InlineData("double", Underline.Double)]
    [InlineData("dotted", Underline.Dotted)]
    [InlineData("dash", Underline.Dash)]
    public void Execute_WithUnderline_AppliesUnderlineStyle(string underlineType, Underline expected)
    {
        var doc = CreateDocumentWithText("Test text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "underline", underlineType }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.Equal(expected, runs[0].Font.Underline);
        AssertModified(context);
    }

    #endregion

    #region Superscript and Subscript

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void Execute_WithSuperscriptOrSubscript_AppliesCorrectly(bool superscript, bool subscript)
    {
        var doc = CreateDocumentWithText("Test text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "superscript", superscript },
            { "subscript", subscript }
        });

        _handler.Execute(context, parameters);

        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.Equal(superscript, runs[0].Font.Superscript);
        Assert.Equal(subscript, runs[0].Font.Subscript);
        AssertModified(context);
    }

    #endregion

    #region Run Index

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Execute_WithRunIndex_FormatsSpecificRun(int runIndex)
    {
        var doc = CreateDocumentWithText("First run");
        var builder = new DocumentBuilder(doc);
        builder.Write(" Second run");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "runIndex", runIndex },
            { "bold", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Run index", result);
        AssertModified(context);
    }

    #endregion

    #region Basic Formatting

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Execute_WithBoldAndItalic_AppliesFormatting(bool bold, bool italic)
    {
        var doc = CreateDocumentWithText("Test text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "bold", bold },
            { "italic", italic }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formatting", result, StringComparison.OrdinalIgnoreCase);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.NotEmpty(runs);
        Assert.Equal(bold, runs[0].Font.Bold);
        Assert.Equal(italic, runs[0].Font.Italic);
        AssertModified(context);
    }

    [Theory]
    [InlineData(8.0)]
    [InlineData(12.0)]
    [InlineData(24.0)]
    [InlineData(72.0)]
    public void Execute_WithFontSize_AppliesSize(double fontSize)
    {
        var doc = CreateDocumentWithText("Test text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "fontSize", fontSize }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formatting", result, StringComparison.OrdinalIgnoreCase);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.Equal(fontSize, runs[0].Font.Size);
        AssertModified(context);
    }

    [Theory]
    [InlineData("Arial")]
    [InlineData("Times New Roman")]
    [InlineData("Calibri")]
    public void Execute_WithFontName_AppliesFont(string fontName)
    {
        var doc = CreateDocumentWithText("Test text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "fontName", fontName }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formatting", result, StringComparison.OrdinalIgnoreCase);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        Assert.Equal(fontName, runs[0].Font.Name);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("paragraphIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(100)]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException(int paragraphIndex)
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", paragraphIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidRunIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "runIndex", 100 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
