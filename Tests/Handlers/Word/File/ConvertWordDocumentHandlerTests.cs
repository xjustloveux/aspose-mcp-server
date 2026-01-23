using Aspose.Words;
using AsposeMcpServer.Handlers.Word.File;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
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

    [Fact]
    public void Execute_WithUnknownExtension_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", Path.Combine(TestDir, "output.abc") }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Cannot infer format", ex.Message);
    }

    #endregion

    #region Additional Format Tests

    [Theory]
    [InlineData("docx")]
    [InlineData("doc")]
    [InlineData("odt")]
    public void Execute_WithDocumentFormats_Converts(string format)
    {
        var outputPath = Path.Combine(TestDir, $"output.{format}");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", outputPath },
            { "format", format }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(System.IO.File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithEpubFormat_Converts()
    {
        var outputPath = Path.Combine(TestDir, "output.epub");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", outputPath },
            { "format", "epub" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(System.IO.File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithXpsFormat_Converts()
    {
        var outputPath = Path.Combine(TestDir, "output.xps");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", outputPath },
            { "format", "xps" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(System.IO.File.Exists(outputPath));
    }

    [Fact]
    public void Execute_InfersPdfFromExtension()
    {
        var outputPath = Path.Combine(TestDir, "inferred.pdf");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("pdf", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_InfersHtmFromExtension()
    {
        var outputPath = Path.Combine(TestDir, "inferred.htm");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("converted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("html", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_CreatesOutputDirectory()
    {
        var subDir = Path.Combine(TestDir, "subdir", "nested");
        var outputPath = Path.Combine(subDir, "output.pdf");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", outputPath },
            { "format", "pdf" }
        });

        _ = _handler.Execute(context, parameters);

        Assert.True(Directory.Exists(subDir));
        Assert.True(System.IO.File.Exists(outputPath));
    }

    #endregion
}
