using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Format;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Format;

public class ClearTabStopsWordHandlerTests : WordHandlerTestBase
{
    private readonly ClearTabStopsWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ClearTabStops()
    {
        Assert.Equal("clear_tab_stops", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTabStops()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with tabs");
        builder.CurrentParagraph.ParagraphFormat.TabStops.Add(new TabStop(72.0, TabAlignment.Left, TabLeader.None));
        builder.CurrentParagraph.ParagraphFormat.TabStops.Add(new TabStop(144.0, TabAlignment.Center, TabLeader.Dots));
        return doc;
    }

    #endregion

    #region Basic Clear Operations

    [Fact]
    public void Execute_ClearsTabStops()
    {
        var doc = CreateDocumentWithTabStops();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("cleared", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoTabStops_ReportsZeroCleared()
    {
        var doc = CreateDocumentWithText("Sample text without tabs.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("cleared", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("0", result.Message);
    }

    #endregion
}
