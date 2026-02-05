using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.FormField;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using PdfForm = Aspose.Pdf.Facades.Form;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FormField;

/// <summary>
///     Unit tests for ImportPdfFormDataHandler class.
/// </summary>
public class ImportPdfFormDataHandlerTests : PdfHandlerTestBase
{
    private readonly ImportPdfFormDataHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Import()
    {
        Assert.Equal("import", _handler.Operation);
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

    private string CreateExportedDataFile(Document doc, string format)
    {
        var exportPath = Path.Combine(TestDir, $"import_test_data.{format}");
        using var form = new PdfForm(doc);
        using var stream = new FileStream(exportPath, FileMode.Create);
        switch (format)
        {
            case "fdf":
                form.ExportFdf(stream);
                break;
            case "xfdf":
                form.ExportXfdf(stream);
                break;
            case "xml":
                form.ExportXml(stream);
                break;
        }

        return exportPath;
    }

    #endregion

    #region Basic Import Operations

    [Fact]
    public void Execute_ImportsXfdfData()
    {
        var sourceDoc = CreateDocumentWithFormFields();
        var dataPath = CreateExportedDataFile(sourceDoc, "xfdf");

        var targetDoc = CreateDocumentWithFormFields();
        (targetDoc.Form["name"] as TextBoxField)!.Value = "";
        (targetDoc.Form["email"] as TextBoxField)!.Value = "";

        var context = CreateContext(targetDoc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataPath", dataPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("XFDF", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Theory]
    [InlineData("fdf", "FDF")]
    [InlineData("xfdf", "XFDF")]
    [InlineData("xml", "XML")]
    public void Execute_WithFormat_ImportsCorrectly(string format, string expectedFormatInMessage)
    {
        var sourceDoc = CreateDocumentWithFormFields();
        var dataPath = CreateExportedDataFile(sourceDoc, format);

        var targetDoc = CreateDocumentWithFormFields();
        var context = CreateContext(targetDoc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataPath", dataPath },
            { "format", format }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains(expectedFormatInMessage, result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_AutoDetectsFormatFromExtension()
    {
        var sourceDoc = CreateDocumentWithFormFields();
        var dataPath = CreateExportedDataFile(sourceDoc, "xfdf");

        var targetDoc = CreateDocumentWithFormFields();
        var context = CreateContext(targetDoc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataPath", dataPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("XFDF", result.Message);
        AssertModified(context);
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
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var doc = CreateDocumentWithFormFields();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataPath", Path.Combine(TestDir, "nonexistent_file.xfdf") }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnknownFormat_ThrowsArgumentException()
    {
        var sourceDoc = CreateDocumentWithFormFields();
        var dataPath = CreateExportedDataFile(sourceDoc, "xfdf");

        var targetDoc = CreateDocumentWithFormFields();
        var context = CreateContext(targetDoc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataPath", dataPath },
            { "format", "unknown" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unknown import format", ex.Message);
    }

    #endregion
}
