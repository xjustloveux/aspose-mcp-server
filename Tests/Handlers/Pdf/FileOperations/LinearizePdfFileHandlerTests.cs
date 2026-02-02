using AsposeMcpServer.Handlers.Pdf.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FileOperations;

public class LinearizePdfFileHandlerTests : PdfHandlerTestBase
{
    private readonly LinearizePdfFileHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Linearize()
    {
        Assert.Equal("linearize", _handler.Operation);
    }

    #endregion

    #region Basic Linearize Operations

    [Fact]
    public void Execute_LinearizesPdf()
    {
        var doc = CreateDocumentWithText("Test content for web viewing");
        var initialPageCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertPageCount(doc, initialPageCount);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_LinearizesMultiPagePdf()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Evaluation mode limits to 4 pages");

        var doc = CreateDocumentWithPages(5);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertPageCount(doc, 5);
        AssertModified(context);
    }

    #endregion
}
