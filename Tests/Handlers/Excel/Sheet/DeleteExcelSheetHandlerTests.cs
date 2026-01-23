using AsposeMcpServer.Handlers.Excel.Sheet;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Sheet;

public class DeleteExcelSheetHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteExcelSheetHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Preserve Other Sheets

    [Fact]
    public void Execute_PreservesOtherSheets()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(2, workbook.Worksheets.Count);
        Assert.NotNull(workbook.Worksheets["Sheet1"]);
        Assert.NotNull(workbook.Worksheets["Sheet3"]);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSheetNameAndIndexInMessage()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("ToDelete");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("ToDelete", result.Message);
        Assert.Contains("1", result.Message);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var initialCount = workbook.Worksheets.Count;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(initialCount - 1, workbook.Worksheets.Count);
        AssertModified(context);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesSheetAtVariousIndices(int sheetIndex)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        workbook.Worksheets.Add("Sheet4");
        var sheetName = workbook.Worksheets[sheetIndex].Name;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", sheetIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(workbook.Worksheets, ws => ws.Name == sheetName);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sheetIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithLastSheet_ThrowsInvalidOperationException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        var ex = Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
        Assert.Contains("last", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidIndex_ThrowsArgumentException(int invalidIndex)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
