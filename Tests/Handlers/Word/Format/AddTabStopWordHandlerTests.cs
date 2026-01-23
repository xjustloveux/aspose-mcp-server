using AsposeMcpServer.Handlers.Word.Format;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Format;

public class AddTabStopWordHandlerTests : WordHandlerTestBase
{
    private readonly AddTabStopWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddTabStop()
    {
        Assert.Equal("add_tab_stop", _handler.Operation);
    }

    #endregion

    #region Various Alignment Types

    [Theory]
    [InlineData("right")]
    [InlineData("decimal")]
    [InlineData("bar")]
    public void Execute_WithVariousAlignments_AddsTabStop(string alignment)
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "tabPosition", 100.0 },
            { "tabAlignment", alignment }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains(alignment, result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Various Leader Types

    [Theory]
    [InlineData("dashes")]
    [InlineData("line")]
    [InlineData("heavy")]
    [InlineData("middledot")]
    public void Execute_WithVariousLeaders_AddsTabStop(string leader)
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "tabPosition", 100.0 },
            { "tabLeader", leader }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains(leader, result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsTabStop()
    {
        var doc = CreateDocumentWithText("Sample text with tab.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "tabPosition", 72.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("tab stop added", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("72", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAlignment_AddsTabStopWithAlignment()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "tabPosition", 144.0 },
            { "tabAlignment", "center" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("center", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithLeader_AddsTabStopWithLeader()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "tabPosition", 200.0 },
            { "tabLeader", "dots" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("dots", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Default Values

    [Fact]
    public void Execute_WithDefaultAlignment_UsesLeft()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "tabPosition", 72.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("left", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithDefaultLeader_UsesNone()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "tabPosition", 72.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("none", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
