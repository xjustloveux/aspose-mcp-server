using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class WordTableHelperTests : WordTestBase
{
    #region GetVerticalAlignment Tests

    [Theory]
    [InlineData("top", CellVerticalAlignment.Top)]
    [InlineData("TOP", CellVerticalAlignment.Top)]
    [InlineData("Top", CellVerticalAlignment.Top)]
    [InlineData("bottom", CellVerticalAlignment.Bottom)]
    [InlineData("BOTTOM", CellVerticalAlignment.Bottom)]
    [InlineData("Bottom", CellVerticalAlignment.Bottom)]
    [InlineData("center", CellVerticalAlignment.Center)]
    [InlineData("CENTER", CellVerticalAlignment.Center)]
    public void GetVerticalAlignment_WithValidValues_ReturnsCorrectAlignment(string input,
        CellVerticalAlignment expected)
    {
        var result = WordTableHelper.GetVerticalAlignment(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    [InlineData("middle")]
    public void GetVerticalAlignment_WithInvalidValues_ReturnsCenter(string input)
    {
        var result = WordTableHelper.GetVerticalAlignment(input);

        Assert.Equal(CellVerticalAlignment.Center, result);
    }

    #endregion

    #region GetLineStyle Tests

    [Theory]
    [InlineData("none", LineStyle.None)]
    [InlineData("NONE", LineStyle.None)]
    [InlineData("single", LineStyle.Single)]
    [InlineData("SINGLE", LineStyle.Single)]
    [InlineData("double", LineStyle.Double)]
    [InlineData("DOUBLE", LineStyle.Double)]
    [InlineData("dotted", LineStyle.Dot)]
    [InlineData("dashed", LineStyle.Single)]
    [InlineData("thick", LineStyle.Thick)]
    [InlineData("THICK", LineStyle.Thick)]
    public void GetLineStyle_WithValidValues_ReturnsCorrectStyle(string input, LineStyle expected)
    {
        var result = WordTableHelper.GetLineStyle(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    public void GetLineStyle_WithInvalidValues_ReturnsSingle(string input)
    {
        var result = WordTableHelper.GetLineStyle(input);

        Assert.Equal(LineStyle.Single, result);
    }

    #endregion

    #region ParseColorDictionary Tests

    [Fact]
    public void ParseColorDictionary_WithNull_ReturnsEmptyDictionary()
    {
        var result = WordTableHelper.ParseColorDictionary(null);

        Assert.Empty(result);
    }

    [Fact]
    public void ParseColorDictionary_WithValidJson_ReturnsDictionary()
    {
        var json = JsonNode.Parse("{\"0\": \"#FF0000\", \"1\": \"#00FF00\", \"2\": \"#0000FF\"}");

        var result = WordTableHelper.ParseColorDictionary(json);

        Assert.Equal(3, result.Count);
        Assert.Equal("#FF0000", result[0]);
        Assert.Equal("#00FF00", result[1]);
        Assert.Equal("#0000FF", result[2]);
    }

    [Fact]
    public void ParseColorDictionary_WithNonNumericKeys_IgnoresInvalidKeys()
    {
        var json = JsonNode.Parse("{\"0\": \"red\", \"invalid\": \"blue\", \"2\": \"green\"}");

        var result = WordTableHelper.ParseColorDictionary(json);

        Assert.Equal(2, result.Count);
        Assert.Equal("red", result[0]);
        Assert.Equal("green", result[2]);
    }

    [Fact]
    public void ParseColorDictionary_WithEmptyObject_ReturnsEmptyDictionary()
    {
        var json = JsonNode.Parse("{}");

        var result = WordTableHelper.ParseColorDictionary(json);

        Assert.Empty(result);
    }

    #endregion

    #region ParseCellColors Tests

    [Fact]
    public void ParseCellColors_WithNull_ReturnsEmptyList()
    {
        var result = WordTableHelper.ParseCellColors(null);

        Assert.Empty(result);
    }

    [Fact]
    public void ParseCellColors_WithValidArray_ReturnsList()
    {
        var json = JsonNode.Parse("[[0, 0, \"#FF0000\"], [1, 2, \"#00FF00\"]]");

        var result = WordTableHelper.ParseCellColors(json);

        Assert.Equal(2, result.Count);
        Assert.Equal((0, 0, "#FF0000"), result[0]);
        Assert.Equal((1, 2, "#00FF00"), result[1]);
    }

    [Fact]
    public void ParseCellColors_WithIncompleteItems_IgnoresInvalidItems()
    {
        var json = JsonNode.Parse("[[0, 0, \"red\"], [1], [2, 3, \"blue\"]]");

        var result = WordTableHelper.ParseCellColors(json);

        Assert.Equal(2, result.Count);
        Assert.Equal((0, 0, "red"), result[0]);
        Assert.Equal((2, 3, "blue"), result[1]);
    }

    [Fact]
    public void ParseCellColors_WithEmptyArray_ReturnsEmptyList()
    {
        var json = JsonNode.Parse("[]");

        var result = WordTableHelper.ParseCellColors(json);

        Assert.Empty(result);
    }

    #endregion

    #region ParseMergeCells Tests

    [Fact]
    public void ParseMergeCells_WithNull_ReturnsEmptyList()
    {
        var result = WordTableHelper.ParseMergeCells(null);

        Assert.Empty(result);
    }

    [Fact]
    public void ParseMergeCells_WithValidArray_ReturnsList()
    {
        var json = JsonNode.Parse(
            "[{\"startRow\": 0, \"endRow\": 1, \"startCol\": 0, \"endCol\": 2}]");

        var result = WordTableHelper.ParseMergeCells(json);

        Assert.Single(result);
        Assert.Equal((0, 1, 0, 2), result[0]);
    }

    [Fact]
    public void ParseMergeCells_WithMultipleItems_ReturnsAll()
    {
        var json = JsonNode.Parse(
            "[{\"startRow\": 0, \"endRow\": 0, \"startCol\": 0, \"endCol\": 1}, " +
            "{\"startRow\": 1, \"endRow\": 2, \"startCol\": 1, \"endCol\": 3}]");

        var result = WordTableHelper.ParseMergeCells(json);

        Assert.Equal(2, result.Count);
        Assert.Equal((0, 0, 0, 1), result[0]);
        Assert.Equal((1, 2, 1, 3), result[1]);
    }

    [Fact]
    public void ParseMergeCells_WithMissingProperties_IgnoresInvalidItems()
    {
        var json = JsonNode.Parse(
            "[{\"startRow\": 0, \"endRow\": 1}, {\"startRow\": 0, \"endRow\": 1, \"startCol\": 0, \"endCol\": 2}]");

        var result = WordTableHelper.ParseMergeCells(json);

        Assert.Single(result);
        Assert.Equal((0, 1, 0, 2), result[0]);
    }

    [Fact]
    public void ParseMergeCells_WithEmptyArray_ReturnsEmptyList()
    {
        var json = JsonNode.Parse("[]");

        var result = WordTableHelper.ParseMergeCells(json);

        Assert.Empty(result);
    }

    #endregion

    #region GetTables Tests

    [Fact]
    public void GetTables_WithoutSectionIndex_ReturnsAllTables()
    {
        var doc = CreateDocumentWithTables(3);

        var result = WordTableHelper.GetTables(doc, null);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void GetTables_WithValidSectionIndex_ReturnsTablesInSection()
    {
        var doc = CreateDocumentWithTableInSection();

        var result = WordTableHelper.GetTables(doc, 0);

        Assert.Single(result);
    }

    [Fact]
    public void GetTables_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTables(1);

        var ex = Assert.Throws<ArgumentException>(() =>
            WordTableHelper.GetTables(doc, 10));

        Assert.Contains("Section index", ex.Message);
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void GetTables_WithNegativeSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTables(1);

        var ex = Assert.Throws<ArgumentException>(() =>
            WordTableHelper.GetTables(doc, -1));

        Assert.Contains("Section index", ex.Message);
    }

    #endregion

    #region GetTable Tests

    [Fact]
    public void GetTable_WithValidIndex_ReturnsTable()
    {
        var doc = CreateDocumentWithTables(3);

        var result = WordTableHelper.GetTable(doc, 1, null);

        Assert.NotNull(result);
    }

    [Fact]
    public void GetTable_WithInvalidIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTables(2);

        var ex = Assert.Throws<ArgumentException>(() =>
            WordTableHelper.GetTable(doc, 5, null));

        Assert.Contains("Table index", ex.Message);
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void GetTable_WithNegativeIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTables(1);

        var ex = Assert.Throws<ArgumentException>(() =>
            WordTableHelper.GetTable(doc, -1, null));

        Assert.Contains("Table index", ex.Message);
    }

    #endregion

    #region ValidateSectionIndex Tests

    [Fact]
    public void ValidateSectionIndex_WithValidIndex_DoesNotThrow()
    {
        var doc = CreateDocumentWithTables(1);

        var exception = Record.Exception(() =>
            WordTableHelper.ValidateSectionIndex(doc, 0));

        Assert.Null(exception);
    }

    [Fact]
    public void ValidateSectionIndex_WithNegativeIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTables(1);

        var ex = Assert.Throws<ArgumentException>(() =>
            WordTableHelper.ValidateSectionIndex(doc, -1));

        Assert.Contains("sectionIndex must be between", ex.Message);
    }

    [Fact]
    public void ValidateSectionIndex_WithIndexTooLarge_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTables(1);

        var ex = Assert.Throws<ArgumentException>(() =>
            WordTableHelper.ValidateSectionIndex(doc, 5));

        Assert.Contains("sectionIndex must be between", ex.Message);
    }

    #endregion

    #region GetPrecedingText Tests

    [Fact]
    public void GetPrecedingText_WithPrecedingParagraph_ReturnsText()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Preceding paragraph text");
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell");
        builder.EndRow();
        builder.EndTable();

        var table = doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().First();

        var result = WordTableHelper.GetPrecedingText(table, 100);

        Assert.Contains("Preceding paragraph text", result);
    }

    [Fact]
    public void GetPrecedingText_WithNoPrecedingContent_ReturnsEmpty()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell");
        builder.EndRow();
        builder.EndTable();

        var table = doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().First();

        var result = WordTableHelper.GetPrecedingText(table, 100);

        Assert.Empty(result);
    }

    [Fact]
    public void GetPrecedingText_WithLongText_TruncatesWithEllipsis()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a very long preceding paragraph text that should be truncated");
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell");
        builder.EndRow();
        builder.EndTable();

        var table = doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().First();

        var result = WordTableHelper.GetPrecedingText(table, 10);

        Assert.EndsWith("...", result);
        Assert.True(result.Length <= 13); // 10 chars + "..."
    }

    #endregion

    #region ApplyMergeCells Tests

    [Fact]
    public void ApplyMergeCells_WithInvalidRange_DoesNothing()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var table = doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().First();

        WordTableHelper.ApplyMergeCells(table, 2, 1, 0, 0);

        Assert.Equal(CellMerge.None, table.Rows[0].Cells[0].CellFormat.VerticalMerge);
    }

    [Fact]
    public void ApplyMergeCells_WithNegativeStartRow_DoesNothing()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var table = doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().First();

        WordTableHelper.ApplyMergeCells(table, -1, 1, 0, 1);

        Assert.Equal(CellMerge.None, table.Rows[0].Cells[0].CellFormat.VerticalMerge);
    }

    [Fact]
    public void ApplyMergeCells_WithRowOutOfRange_DoesNothing()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var table = doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().First();

        WordTableHelper.ApplyMergeCells(table, 0, 10, 0, 1);

        Assert.Equal(CellMerge.None, table.Rows[0].Cells[0].CellFormat.VerticalMerge);
    }

    [Fact]
    public void ApplyMergeCells_WithValidHorizontalMerge_SetsHorizontalMerge()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var table = doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().First();

        WordTableHelper.ApplyMergeCells(table, 0, 0, 0, 1);

        Assert.Equal(CellMerge.First, table.Rows[0].Cells[0].CellFormat.HorizontalMerge);
        Assert.Equal(CellMerge.Previous, table.Rows[0].Cells[1].CellFormat.HorizontalMerge);
    }

    [Fact]
    public void ApplyMergeCells_WithValidVerticalMerge_SetsVerticalMerge()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var table = doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().First();

        WordTableHelper.ApplyMergeCells(table, 0, 1, 0, 0);

        Assert.Equal(CellMerge.First, table.Rows[0].Cells[0].CellFormat.VerticalMerge);
        Assert.Equal(CellMerge.Previous, table.Rows[1].Cells[0].CellFormat.VerticalMerge);
    }

    [Fact]
    public void ApplyMergeCells_WithBothMergeDirections_SetsBothMerges()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var table = doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().First();

        WordTableHelper.ApplyMergeCells(table, 0, 1, 0, 1);

        Assert.Equal(CellMerge.First, table.Rows[0].Cells[0].CellFormat.VerticalMerge);
        Assert.Equal(CellMerge.First, table.Rows[0].Cells[0].CellFormat.HorizontalMerge);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTables(int tableCount)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        for (var i = 0; i < tableCount; i++)
        {
            builder.StartTable();
            builder.InsertCell();
            builder.Write($"Table {i + 1}");
            builder.EndRow();
            builder.EndTable();
            builder.Writeln();
        }

        return doc;
    }

    private static Document CreateDocumentWithTableInSection()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell in section 0");
        builder.EndRow();
        builder.EndTable();

        return doc;
    }

    private static Document CreateDocumentWithTable(int rows, int cols)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartTable();

        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r}C{c}");
            }

            builder.EndRow();
        }

        builder.EndTable();
        return doc;
    }

    #endregion
}
