using AsposeMcpServer.Handlers.Excel.Protect;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Protect;

public class ProtectExcelHandlerTests : ExcelHandlerTestBase
{
    private readonly ProtectExcelHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Protect()
    {
        Assert.Equal("protect", _handler.Operation);
    }

    #endregion

    #region Workbook Protection

    [Fact]
    public void Execute_ProtectsWorkbook()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" },
            { "protectWorkbook", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("protected", result.ToLower());
        Assert.Contains("workbook", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithProtectStructure_ProtectsStructure()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" },
            { "protectWorkbook", true },
            { "protectStructure", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("protected", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithProtectWindows_ProtectsWindows()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" },
            { "protectWorkbook", true },
            { "protectWindows", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("protected", result.ToLower());
        AssertModified(context);
    }

    #endregion

    #region Worksheet Protection

    [Fact]
    public void Execute_WithSheetIndex_ProtectsWorksheet()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" },
            { "sheetIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("protected", result.ToLower());
        Assert.Contains("worksheet", result.ToLower());
        Assert.True(workbook.Worksheets[0].IsProtected);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithMultipleSheets_ProtectsCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" },
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.False(workbook.Worksheets[0].IsProtected);
        Assert.True(workbook.Worksheets[1].IsProtected);
        Assert.False(workbook.Worksheets[2].IsProtected);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPassword_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyPassword_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("password", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" },
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
