using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Protect;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Protect;

public class UnprotectExcelHandlerTests : ExcelHandlerTestBase
{
    private readonly UnprotectExcelHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Unprotect()
    {
        Assert.Equal("unprotect", _handler.Operation);
    }

    #endregion

    #region Workbook Unprotection

    [Fact]
    public void Execute_UnprotectsWorkbook()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Protect(ProtectionType.Structure, "test123");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "password", "test123" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("removed", result.ToLower());
        Assert.Contains("workbook", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_UnprotectsWorkbookWithoutPassword()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("workbook", result.ToLower());
        AssertModified(context);
    }

    #endregion

    #region Worksheet Unprotection

    [Fact]
    public void Execute_WithSheetIndex_UnprotectsWorksheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);
        Assert.True(workbook.Worksheets[0].IsProtected);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "password", "test123" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("removed", result.ToLower());
        Assert.False(workbook.Worksheets[0].IsProtected);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithUnprotectedSheet_ReturnsNotProtectedMessage()
    {
        var workbook = CreateEmptyWorkbook();
        Assert.False(workbook.Worksheets[0].IsProtected);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("not protected", result.ToLower());
    }

    [Fact]
    public void Execute_WithMultipleSheets_UnprotectsCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        workbook.Worksheets[0].Protect(ProtectionType.All, "pass0", null);
        workbook.Worksheets[1].Protect(ProtectionType.All, "pass1", null);
        workbook.Worksheets[2].Protect(ProtectionType.All, "pass2", null);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "password", "pass1" }
        });

        _handler.Execute(context, parameters);

        Assert.True(workbook.Worksheets[0].IsProtected);
        Assert.False(workbook.Worksheets[1].IsProtected);
        Assert.True(workbook.Worksheets[2].IsProtected);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithWrongPassword_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Protect(ProtectionType.All, "correct", null);

        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "password", "wrong" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("password", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
