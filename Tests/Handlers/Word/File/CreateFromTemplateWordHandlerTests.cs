using Aspose.Words;
using AsposeMcpServer.Handlers.Word.File;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.File;

public class CreateFromTemplateWordHandlerTests : WordHandlerTestBase
{
    private readonly CreateFromTemplateWordHandler _handler = new();
    private readonly string _templatePath;

    public CreateFromTemplateWordHandlerTests()
    {
        _templatePath = Path.Combine(TestDir, "template.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Hello <<[name]>>!");
        doc.Save(_templatePath);
    }

    #region Operation Property

    [Fact]
    public void Operation_Returns_CreateFromTemplate()
    {
        Assert.Equal("create_from_template", _handler.Operation);
    }

    #endregion

    #region Basic Create Operations

    [Fact]
    public void Execute_CreatesDocumentFromTemplate()
    {
        var outputPath = Path.Combine(TestDir, "output.docx");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "templatePath", _templatePath },
            { "outputPath", outputPath },
            { "dataJson", "{\"name\": \"World\"}" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created from template", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(System.IO.File.Exists(outputPath));
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutTemplatePathOrSessionId_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", Path.Combine(TestDir, "output.docx") },
            { "dataJson", "{\"name\": \"World\"}" }
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
            { "templatePath", _templatePath },
            { "dataJson", "{\"name\": \"World\"}" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutDataJson_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "templatePath", _templatePath },
            { "outputPath", Path.Combine(TestDir, "output.docx") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentTemplate_ThrowsFileNotFoundException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "templatePath", Path.Combine(TestDir, "nonexistent.docx") },
            { "outputPath", Path.Combine(TestDir, "output.docx") },
            { "dataJson", "{\"name\": \"World\"}" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
