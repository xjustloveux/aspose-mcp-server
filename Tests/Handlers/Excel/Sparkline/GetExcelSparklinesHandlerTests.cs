using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Handlers.Excel.Sparkline;
using AsposeMcpServer.Results.Excel.Sparkline;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Sparkline;

public class GetExcelSparklinesHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelSparklinesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeGet()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Execute Tests

    [Fact]
    public void Execute_WithNoSparklines_ShouldReturnEmpty()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSparklinesExcelResult>(res);
        Assert.Equal(0, result.Count);
        Assert.Empty(result.Items);
        Assert.NotNull(result.Message);
        Assert.Contains("No sparkline groups", result.Message);
    }

    [SkippableFact]
    public void Execute_WithSparkline_ShouldReturnInfo()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 1 }, { 2 }, { 3 }, { 4 }, { 5 }
        });
        var worksheet = workbook.Worksheets[0];
        var sheetName = worksheet.Name;
        var locationArea = CellArea.CreateCellArea("B1", "B1");
        worksheet.SparklineGroups.Add(SparklineType.Line, $"{sheetName}!A1:A5", true, locationArea);

        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSparklinesExcelResult>(res);
        Assert.True(result.Count > 0);
        Assert.NotEmpty(result.Items);

        var item = result.Items[0];
        Assert.Equal(0, item.Index);
        Assert.NotNull(item.Type);
    }

    #endregion
}
