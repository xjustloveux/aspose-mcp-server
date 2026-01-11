using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class CreateWordTableHandlerTests : WordHandlerTestBase
{
    private readonly CreateWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Create()
    {
        Assert.Equal("create", _handler.Operation);
    }

    #endregion

    #region Basic Create Operations

    [Fact]
    public void Execute_CreatesTable()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("created", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(GetTableCount(doc) > 0);
        AssertModified(context);
    }

    [Theory]
    [InlineData(2, 3)]
    [InlineData(4, 5)]
    [InlineData(10, 10)]
    public void Execute_WithRowsAndColumns_CreatesTableWithSpecifiedSize(int rows, int cols)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rows", rows },
            { "columns", cols }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"{rows} rows", result);
        Assert.Contains($"{cols} columns", result);
        var table = GetFirstTable(doc);
        Assert.Equal(rows, table.Rows.Count);
        Assert.Equal(cols, table.Rows[0].Cells.Count);
    }

    [Fact]
    public void Execute_Default_Creates3x3Table()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        var table = GetFirstTable(doc);
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(3, table.Rows[0].Cells.Count);
    }

    #endregion

    #region Table Data

    [Fact]
    public void Execute_WithTableData_CreatesTableWithData()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableData", "[[\"A1\", \"B1\"], [\"A2\", \"B2\"]]" }
        });

        _handler.Execute(context, parameters);

        var table = GetFirstTable(doc);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(2, table.Rows[0].Cells.Count);
        AssertContainsText(doc, "A1");
        AssertContainsText(doc, "B2");
    }

    [Fact]
    public void Execute_WithInvalidTableData_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableData", "invalid json" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Table Options

    [Fact]
    public void Execute_WithTableWidth_SetsWidth()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableWidth", 400.0 }
        });

        _handler.Execute(context, parameters);

        var table = GetFirstTable(doc);
        Assert.Equal(PreferredWidthType.Points, table.PreferredWidth.Type);
        Assert.Equal(400.0, table.PreferredWidth.Value, 1);
    }

    [Fact]
    public void Execute_WithAutoFitFalse_DisablesAutoFit()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "autoFit", false }
        });

        _handler.Execute(context, parameters);

        var table = GetFirstTable(doc);
        Assert.False(table.AllowAutoFit);
    }

    [Fact]
    public void Execute_WithHasHeaderTrue_MakesFirstRowBold()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableData", "[[\"Header1\", \"Header2\"], [\"Data1\", \"Data2\"]]" },
            { "hasHeader", true }
        });

        _handler.Execute(context, parameters);

        var table = GetFirstTable(doc);
        var firstCellRuns = table.Rows[0].Cells[0].GetChildNodes(NodeType.Run, true);
        if (firstCellRuns.Count > 0)
        {
            var run = (Run)firstCellRuns[0];
            Assert.True(run.Font.Bold);
        }
    }

    #endregion

    #region Section Index

    [Fact]
    public void Execute_WithSectionIndex_CreatesInSpecificSection()
    {
        var doc = CreateDocumentWithSections(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 }
        });

        _handler.Execute(context, parameters);

        var tablesInSection1 = doc.Sections[1].Body.GetChildNodes(NodeType.Table, true);
        Assert.True(tablesInSection1.Count > 0);
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sectionIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Helper Methods

    private static int GetTableCount(Document doc)
    {
        return doc.GetChildNodes(NodeType.Table, true).Count;
    }

    private static Aspose.Words.Tables.Table GetFirstTable(Document doc)
    {
        return (Aspose.Words.Tables.Table)doc.GetChildNodes(NodeType.Table, true)[0];
    }

    private static Document CreateDocumentWithSections(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 1; i < count; i++) builder.InsertBreak(BreakType.SectionBreakNewPage);
        return doc;
    }

    #endregion
}
