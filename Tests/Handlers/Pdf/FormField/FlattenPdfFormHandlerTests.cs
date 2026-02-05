using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.FormField;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FormField;

/// <summary>
///     Unit tests for FlattenPdfFormHandler class.
/// </summary>
public class FlattenPdfFormHandlerTests : PdfHandlerTestBase
{
    private readonly FlattenPdfFormHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Flatten()
    {
        Assert.Equal("flatten", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithFormFields(int fieldCount = 3)
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        for (var i = 0; i < fieldCount; i++)
        {
            var field = new TextBoxField(page, new Rectangle(100, 700 - i * 30, 300, 720 - i * 30))
            {
                PartialName = $"field{i + 1}",
                Value = $"Value {i + 1}"
            };
            doc.Form.Add(field);
        }

        return doc;
    }

    #endregion

    #region Result Message

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(5)]
    public void Execute_ReturnsMessageWithFieldCount(int fieldCount)
    {
        var doc = CreateDocumentWithFormFields(fieldCount);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains(fieldCount.ToString(), result.Message);
        Assert.Contains("static content", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Flatten Operations

    [Fact]
    public void Execute_FlattensFormFields()
    {
        var doc = CreateDocumentWithFormFields();
        Assert.Equal(3, doc.Form.Count);

        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("3", result.Message);
        Assert.Contains("flatten", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSingleField_FlattensSuccessfully()
    {
        var doc = CreateDocumentWithFormFields(1);
        Assert.Single(doc.Form);

        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("1", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_OnDocumentWithNoFields_StillSucceeds()
    {
        var doc = CreateDocumentWithPages(1);
        Assert.Empty(doc.Form);

        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("0", result.Message);
        AssertModified(context);
    }

    #endregion
}
