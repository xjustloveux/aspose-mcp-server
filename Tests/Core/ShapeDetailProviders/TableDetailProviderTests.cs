using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

public class TableDetailProviderTests : TestBase
{
    private readonly TableDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnTable()
    {
        Assert.Equal("Table", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithTable_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50], [30]);

        var result = _provider.CanHandle(table);

        Assert.True(result);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.False(result);
    }

    [Fact]
    public void GetDetails_WithTable_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);

        var details = _provider.GetDetails(table, presentation);

        Assert.NotNull(details);
    }

    [Fact]
    public void GetDetails_WithNonTable_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.Null(details);
    }

    [Fact]
    public void GetDetails_WithTable_ShouldIncludeRowAndColumnCount()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50, 50], [30, 30]);

        var details = _provider.GetDetails(table, presentation);
        Assert.NotNull(details);

        var json = JsonSerializer.Serialize(details);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.Equal(2, root.GetProperty("rows").GetInt32());
        Assert.Equal(3, root.GetProperty("columns").GetInt32());
        Assert.Equal(6, root.GetProperty("totalCells").GetInt32());
    }

    [Fact]
    public void GetDetails_WithMergedCells_ShouldIncludeMergedCellInfo()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50, 50], [30, 30, 30]);

        table.MergeCells(table[0, 0], table[1, 0], false);

        var details = _provider.GetDetails(table, presentation);
        Assert.NotNull(details);

        var json = JsonSerializer.Serialize(details);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.True(root.GetProperty("mergedCellCount").GetInt32() > 0);
        Assert.True(root.TryGetProperty("mergedCells", out var mergedCells));
        Assert.NotEqual(JsonValueKind.Null, mergedCells.ValueKind);
    }

    [Fact]
    public void GetDetails_WithNoMergedCells_ShouldHaveNullMergedCells()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);

        var details = _provider.GetDetails(table, presentation);
        Assert.NotNull(details);

        var json = JsonSerializer.Serialize(details);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.Equal(0, root.GetProperty("mergedCellCount").GetInt32());
        Assert.Equal(JsonValueKind.Null, root.GetProperty("mergedCells").ValueKind);
    }
}
