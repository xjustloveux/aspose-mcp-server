using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Handlers.Excel.Sparkline;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Sparkline;

public class DeleteExcelSparklineHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteExcelSparklineHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeDelete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Execute Tests

    [Fact]
    public void Execute_WithMissingGroupIndex_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("groupIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidGroupIndex_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "groupIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_DeletesOnlyTheIndexedGroup_NotOverlappingNeighbors()
    {
        var workbook = CreateEmptyWorkbook();
        var sheet = workbook.Worksheets[0];
        for (var i = 1; i <= 5; i++) sheet.Cells[$"A{i}"].Value = i;

        // Group 0 spans D1 and D3; group 1 sits at D2, inside group 0's bounding box —
        // an area-based clear of group 0 would wrongly take group 1 with it.
        var g0 = sheet.SparklineGroups.Add(SparklineType.Line, $"{sheet.Name}!A1:A5", true,
            CellArea.CreateCellArea("D1", "D1"));
        sheet.SparklineGroups[g0].Sparklines.Add($"{sheet.Name}!A1:A5", 2, 3);
        sheet.SparklineGroups.Add(SparklineType.Line, $"{sheet.Name}!A1:A5", true,
            CellArea.CreateCellArea("D2", "D2"));
        Assert.Equal(2, sheet.SparklineGroups.Count);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "groupIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Single(sheet.SparklineGroups);
        Assert.True(sheet.SparklineGroups[0].Sparklines.Count > 0,
            "The surviving group must keep its sparklines");
    }

    #endregion
}
