using AsposeMcpServer.Handlers.Excel.Sheet;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Sheet;

public class AddExcelSheetHandlerTests : ExcelHandlerTestBase
{
    private readonly AddExcelSheetHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSuccessMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetName", "TestSheet" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("TestSheet", result.Message);
        Assert.Contains("added", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsNewSheet()
    {
        var workbook = CreateEmptyWorkbook();
        var initialCount = workbook.Worksheets.Count;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetName", "NewSheet" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("NewSheet", result.Message);
        Assert.Equal(initialCount + 1, workbook.Worksheets.Count);
        AssertModified(context);
    }

    [Theory]
    [InlineData("Sheet2")]
    [InlineData("Data")]
    [InlineData("Report")]
    public void Execute_AddsSheetWithVariousNames(string sheetName)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetName", sheetName }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains(sheetName, result.Message);
        Assert.NotNull(workbook.Worksheets[sheetName]);
        AssertModified(context);
    }

    #endregion

    #region Insert Position

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithInsertAt_InsertsAtCorrectPosition(int insertAt)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetName", "NewSheet" },
            { "insertAt", insertAt }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(3, workbook.Worksheets.Count);
        Assert.NotNull(workbook.Worksheets["NewSheet"]);
    }

    [Fact]
    public void Execute_WithoutInsertAt_AppendsAtEnd()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetName", "LastSheet" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("LastSheet", workbook.Worksheets[^1].Name);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSheetName_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sheetName", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithDuplicateName_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetName", "Sheet1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("already exists", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyName_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetName", "" }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(100)]
    public void Execute_WithInvalidInsertAt_ThrowsArgumentException(int invalidInsertAt)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetName", "NewSheet" },
            { "insertAt", invalidInsertAt }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("insertAt", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
