using Aspose.Words;
using AsposeMcpServer.Handlers.Word.File;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.File;

public class ConvertWordDocumentHandlerTests : WordHandlerTestBase
{
    private readonly ConvertWordDocumentHandler _handler = new();
    private readonly string _inputPath;

    public ConvertWordDocumentHandlerTests()
    {
        _inputPath = Path.Combine(TestDir, "input.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content for conversion");
        doc.Save(_inputPath);
    }

    #region Operation Property

    [Fact]
    public void Operation_Returns_Convert()
    {
        Assert.Equal("convert", _handler.Operation);
    }

    #endregion

    #region Basic Convert Operations

    [Fact]
    public void Execute_ConvertsToPdf()
    {
        var outputPath = Path.Combine(TestDir, "output.pdf");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", outputPath },
            { "format", "pdf" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("converted", result.ToLower());
        Assert.True(System.IO.File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted PDF file should have content");
    }

    [Fact]
    public void Execute_ConvertsToHtml()
    {
        var outputPath = Path.Combine(TestDir, "output.html");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", outputPath },
            { "format", "html" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("converted", result.ToLower());
        Assert.True(System.IO.File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted HTML file should have content");

        var htmlContent = System.IO.File.ReadAllText(outputPath);
        Assert.Contains("Test content for conversion", htmlContent);
    }

    [Fact]
    public void Execute_ConvertsToRtf()
    {
        var outputPath = Path.Combine(TestDir, "output.rtf");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", outputPath },
            { "format", "rtf" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("converted", result.ToLower());
        Assert.True(System.IO.File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted RTF file should have content");
    }

    [Fact]
    public void Execute_InfersFormatFromExtension()
    {
        var outputPath = Path.Combine(TestDir, "output.txt");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("converted", result.ToLower());
        Assert.True(System.IO.File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted TXT file should have content");

        var txtContent = System.IO.File.ReadAllText(outputPath);
        Assert.Contains("Test content for conversion", txtContent);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPathOrSessionId_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", Path.Combine(TestDir, "output.pdf") },
            { "format", "pdf" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "format", "pdf" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedFormat_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", Path.Combine(TestDir, "output.xyz") },
            { "format", "xyz" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
