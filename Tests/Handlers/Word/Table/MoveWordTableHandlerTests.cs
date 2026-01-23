using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class MoveWordTableHandlerTests : WordHandlerTestBase
{
    private readonly MoveWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_MoveTable()
    {
        Assert.Equal("move_table", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTableAndParagraphs()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");

        builder.StartTable();
        for (var i = 0; i < 2; i++)
        {
            for (var j = 0; j < 2; j++)
            {
                builder.InsertCell();
                builder.Write($"R{i}C{j}");
            }

            builder.EndRow();
        }

        builder.EndTable();

        return doc;
    }

    #endregion

    #region Basic Move Operations

    [Fact]
    public void Execute_MovesTable()
    {
        var doc = CreateDocumentWithTableAndParagraphs();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetParagraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("moved", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_MovesToVariousParagraphs(int paragraphIndex)
    {
        var doc = CreateDocumentWithTableAndParagraphs();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetParagraphIndex", paragraphIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("moved", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidTableIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTableAndParagraphs();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("tableIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTableAndParagraphs();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetParagraphIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("targetParagraphIndex", ex.Message);
    }

    #endregion
}
