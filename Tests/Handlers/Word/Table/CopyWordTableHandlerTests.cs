using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class CopyWordTableHandlerTests : WordHandlerTestBase
{
    private readonly CopyWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_CopyTable()
    {
        Assert.Equal("copy_table", _handler.Operation);
    }

    #endregion

    #region Basic Copy Operations

    [Fact]
    public void Execute_CopiesTable()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("copied", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(2, GetTableCount(doc));
        AssertModified(context);
    }

    [Fact]
    public void Execute_CopiesTableStructure()
    {
        var doc = CreateDocumentWithTable(4, 5);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        var tables = GetTables(doc);
        Assert.Equal(tables[0].Rows.Count, tables[1].Rows.Count);
        Assert.Equal(tables[0].Rows[0].Cells.Count, tables[1].Rows[0].Cells.Count);
    }

    #endregion

    #region Table Index

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Execute_WithTableIndex_CopiesSpecificTable(int tableIndex)
    {
        var doc = CreateDocumentWithTables(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", tableIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("copied", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(4, GetTableCount(doc));
    }

    [Fact]
    public void Execute_WithInvalidTableIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sourceTableIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeTableIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sourceTableIndex", ex.Message);
    }

    #endregion

    #region Target Paragraph Index

    [Fact]
    public void Execute_WithTargetParagraphIndex_CopiesToSpecifiedLocation()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetParagraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("copied", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidTargetParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetParagraphIndex", 999 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("targetParagraphIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeTargetParagraphIndex_CopiesToEnd()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetParagraphIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("copied", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(2, GetTableCount(doc));
    }

    #endregion

    #region Section Index

    [Fact]
    public void Execute_WithInvalidSourceSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceSectionIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sourceSectionIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidTargetSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetSectionIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("targetSectionIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeSourceSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceSectionIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sourceSectionIndex", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static int GetTableCount(Document doc)
    {
        return doc.GetChildNodes(NodeType.Table, true).Count;
    }

    private static List<Aspose.Words.Tables.Table> GetTables(Document doc)
    {
        return doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
    }

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

    #endregion
}
