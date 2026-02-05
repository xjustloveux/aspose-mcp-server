using Aspose.Cells.Charts;
using AsposeMcpServer.Handlers.Excel.Sparkline;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Sparkline;

public class AddExcelSparklineHandlerTests : ExcelHandlerTestBase
{
    private readonly AddExcelSparklineHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeAdd()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Execute Tests

    [SkippableFact]
    public void Execute_WithValidParams_ShouldAddSparkline()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 1 }, { 2 }, { 3 }, { 4 }, { 5 }
        });
        var sheetName = workbook.Worksheets[0].Name;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", $"{sheetName}!A1:A5" },
            { "locationRange", "B1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Sparkline group", result.Message);
        Assert.Contains("A1:A5", result.Message);
        Assert.Contains("B1", result.Message);
    }

    [SkippableFact]
    public void Execute_ShouldMarkModified()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Sparkline operations have limitations in evaluation mode");

        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 1 }, { 2 }, { 3 }, { 4 }, { 5 }
        });
        var sheetName = workbook.Worksheets[0].Name;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", $"{sheetName}!A1:A5" },
            { "locationRange", "B1" }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithMissingDataRange_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "locationRange", "B1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("dataRange", ex.Message);
    }

    [Fact]
    public void Execute_WithMissingLocationRange_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dataRange", "A1:A5" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("locationRange", ex.Message);
    }

    #endregion

    #region ResolveSparklineType Tests

    [Theory]
    [InlineData("line", SparklineType.Line)]
    [InlineData("column", SparklineType.Column)]
    [InlineData("stacked", SparklineType.Stacked)]
    public void ResolveSparklineType_WithValidTypes_ShouldReturn(string type, SparklineType expected)
    {
        var result = AddExcelSparklineHandler.ResolveSparklineType(type);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void ResolveSparklineType_WithInvalidType_ShouldThrow()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            AddExcelSparklineHandler.ResolveSparklineType("invalid"));
        Assert.Contains("Unknown sparkline type", ex.Message);
    }

    #endregion
}
