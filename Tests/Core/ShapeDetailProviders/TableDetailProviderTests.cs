using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;
using AsposeMcpServer.Core.ShapeDetailProviders.Providers;
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
        var tableDetails = Assert.IsType<TableDetails>(details);

        Assert.Equal(2, tableDetails.Rows);
        Assert.Equal(2, tableDetails.Columns);
        Assert.Equal(4, tableDetails.TotalCells);
        Assert.Equal(0, tableDetails.MergedCellCount);
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
        var tableDetails = Assert.IsType<TableDetails>(details);

        Assert.Equal(2, tableDetails.Rows);
        Assert.Equal(3, tableDetails.Columns);
        Assert.Equal(6, tableDetails.TotalCells);
    }

    [Fact]
    public void GetDetails_WithMergedCells_ShouldIncludeMergedCellInfo()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50, 50], [30, 30, 30]);

        table.MergeCells(table[0, 0], table[1, 0], false);

        var details = _provider.GetDetails(table, presentation);
        var tableDetails = Assert.IsType<TableDetails>(details);

        Assert.True(tableDetails.MergedCellCount > 0);
        Assert.NotNull(tableDetails.MergedCells);
    }

    [Fact]
    public void GetDetails_WithNoMergedCells_ShouldHaveNullMergedCells()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);

        var details = _provider.GetDetails(table, presentation);
        var tableDetails = Assert.IsType<TableDetails>(details);

        Assert.Equal(0, tableDetails.MergedCellCount);
        Assert.Null(tableDetails.MergedCells);
    }

    [Fact]
    public void GetDetails_WithTable_ShouldIncludeStyleFlags()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);
        table.FirstRow = true;
        table.FirstCol = true;
        table.LastRow = false;
        table.LastCol = false;
        var details = _provider.GetDetails(table, presentation);
        var tableDetails = Assert.IsType<TableDetails>(details);

        Assert.True(tableDetails.FirstRow);
        Assert.True(tableDetails.FirstCol);
        Assert.False(tableDetails.LastRow);
        Assert.False(tableDetails.LastCol);
    }

    [Fact]
    public void GetDetails_WithTable_ShouldIncludeStylePreset()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);

        var details = _provider.GetDetails(table, presentation);
        var tableDetails = Assert.IsType<TableDetails>(details);

        // StylePreset is either a valid preset name or null (when "None")
        Assert.True(tableDetails.StylePreset == null || !string.IsNullOrEmpty(tableDetails.StylePreset));
    }
}
