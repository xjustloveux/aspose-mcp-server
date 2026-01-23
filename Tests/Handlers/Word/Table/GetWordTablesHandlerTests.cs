using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Results.Word.Table;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class GetWordTablesHandlerTests : WordHandlerTestBase
{
    private readonly GetWordTablesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Table Properties

    [Fact]
    public void Execute_ReturnsTableProperties()
    {
        var doc = CreateDocumentWithTable(4, 5);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTablesWordResult>(res);

        Assert.True(result.Tables.Count > 0);
        var firstTable = result.Tables[0];
        Assert.Equal(0, firstTable.Index);
        Assert.Equal(4, firstTable.Rows);
        Assert.Equal(5, firstTable.Columns);
    }

    #endregion

    #region Section Index

    [Fact]
    public void Execute_WithSectionIndex_GetsFromSpecificSection()
    {
        var doc = CreateDocumentWithSections(2);
        AddTableToSection(doc, 1, 2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTablesWordResult>(res);

        Assert.Equal(1, result.Count);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var initialCount = doc.GetChildNodes(NodeType.Table, true).Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, doc.GetChildNodes(NodeType.Table, true).Count);
        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsTablesInfo()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTablesWordResult>(res);

        Assert.True(result.Count >= 0);
        Assert.NotNull(result.Tables);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithTables(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTablesWordResult>(res);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void Execute_WithNoTables_ReturnsEmptyList()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTablesWordResult>(res);

        Assert.Equal(0, result.Count);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTable(int rows, int cols)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        for (var i = 0; i < rows; i++)
        {
            for (var j = 0; j < cols; j++)
            {
                builder.InsertCell();
                builder.Write($"R{i}C{j}");
            }

            builder.EndRow();
        }

        builder.EndTable();
        return doc;
    }

    private static Document CreateDocumentWithTables(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var t = 0; t < count; t++)
        {
            builder.StartTable();
            for (var i = 0; i < 2; i++)
            {
                for (var j = 0; j < 2; j++)
                {
                    builder.InsertCell();
                    builder.Write($"T{t}R{i}C{j}");
                }

                builder.EndRow();
            }

            builder.EndTable();
            builder.Writeln();
        }

        return doc;
    }

    private static Document CreateDocumentWithSections(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 1; i < count; i++) builder.InsertBreak(BreakType.SectionBreakNewPage);
        return doc;
    }

    private static void AddTableToSection(Document doc, int sectionIndex, int rows, int cols)
    {
        var builder = new DocumentBuilder(doc);
        builder.MoveToSection(sectionIndex);
        builder.StartTable();
        for (var i = 0; i < rows; i++)
        {
            for (var j = 0; j < cols; j++)
            {
                builder.InsertCell();
                builder.Write($"S{sectionIndex}R{i}C{j}");
            }

            builder.EndRow();
        }

        builder.EndTable();
    }

    #endregion
}
