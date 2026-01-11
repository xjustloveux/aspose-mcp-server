using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class GetStructureWordTableHandlerTests : WordHandlerTestBase
{
    private readonly GetStructureWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetStructure()
    {
        Assert.Equal("get_structure", _handler.Operation);
    }

    #endregion

    #region Table Format

    [Fact]
    public void Execute_ReturnsTableFormat()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("[Table Format]", result);
        Assert.Contains("Alignment:", result);
        Assert.Contains("Allow Auto Fit:", result);
    }

    #endregion

    #region Cell Formatting

    [Fact]
    public void Execute_WithIncludeCellFormatting_ReturnsCellFormatting()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeCellFormatting", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("[First Cell Formatting]", result);
        Assert.Contains("Padding:", result);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Structure Operations

    [Fact]
    public void Execute_ReturnsStructureInfo()
    {
        var doc = CreateDocumentWithTable(3, 4);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Structure", result);
        Assert.Contains("Rows: 3", result);
        Assert.Contains("Columns: 4", result);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsBasicInfo()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("[Basic Info]", result);
        Assert.Contains("Rows:", result);
        Assert.Contains("Columns:", result);
    }

    #endregion

    #region Include Content

    [Fact]
    public void Execute_WithIncludeContent_ReturnsContentPreview()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeContent", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("[Content Preview]", result);
        Assert.Contains("Row 0:", result);
    }

    [Fact]
    public void Execute_WithoutIncludeContent_DoesNotReturnContentPreview()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeContent", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.DoesNotContain("[Content Preview]", result);
    }

    #endregion

    #region Table Index

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Execute_WithTableIndex_GetsSpecificTable(int tableIndex)
    {
        var doc = CreateDocumentWithTables(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", tableIndex }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"Table #{tableIndex}", result);
    }

    [Fact]
    public void Execute_WithInvalidTableIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Table index", ex.Message);
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

    #endregion
}
