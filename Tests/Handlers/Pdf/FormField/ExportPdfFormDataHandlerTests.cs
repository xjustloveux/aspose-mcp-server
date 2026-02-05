using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.FormField;
using AsposeMcpServer.Results.Pdf.FormField;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FormField;

/// <summary>
///     Unit tests for ExportPdfFormDataHandler class.
/// </summary>
public class ExportPdfFormDataHandlerTests : PdfHandlerTestBase
{
    private readonly ExportPdfFormDataHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Export()
    {
        Assert.Equal("export", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithFormFields()
    {
        var doc = new Document();
        var page = doc.Pages.Add();

        var field1 = new TextBoxField(page, new Rectangle(100, 700, 300, 720))
        {
            PartialName = "name",
            Value = "John Doe"
        };
        doc.Form.Add(field1);

        var field2 = new TextBoxField(page, new Rectangle(100, 660, 300, 680))
        {
            PartialName = "email",
            Value = "john@example.com"
        };
        doc.Form.Add(field2);

        return doc;
    }

    #endregion

    #region Basic Export Operations

    [Fact]
    public void Execute_ExportsToXfdf_DefaultFormat()
    {
        var doc = CreateDocumentWithFormFields();
        var context = CreateContext(doc);
        var exportPath = Path.Combine(TestDir, "export_default.xfdf");
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataPath", exportPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExportFormDataResult>(res);
        Assert.Contains("XFDF", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("xfdf", result.Format);
        Assert.Equal(exportPath, result.ExportPath);
        Assert.True(File.Exists(exportPath));
        Assert.True(new FileInfo(exportPath).Length > 0);
    }

    [Theory]
    [InlineData("fdf", "FDF")]
    [InlineData("xfdf", "XFDF")]
    [InlineData("xml", "XML")]
    public void Execute_WithFormat_ExportsInCorrectFormat(string format, string expectedFormatInMessage)
    {
        var doc = CreateDocumentWithFormFields();
        var context = CreateContext(doc);
        var exportPath = Path.Combine(TestDir, $"export_test.{format}");
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataPath", exportPath },
            { "format", format }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExportFormDataResult>(res);
        Assert.Contains(expectedFormatInMessage, result.Message);
        Assert.Equal(format, result.Format);
        Assert.True(File.Exists(exportPath));
    }

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithFormFields();
        var context = CreateContext(doc);
        var exportPath = Path.Combine(TestDir, "export_nomod.xfdf");
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataPath", exportPath }
        });

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutDataPath_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithFormFields();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnknownFormat_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithFormFields();
        var context = CreateContext(doc);
        var exportPath = Path.Combine(TestDir, "export_unknown.dat");
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataPath", exportPath },
            { "format", "unknown" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unknown export format", ex.Message);
    }

    #endregion
}
