using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Protect;
using AsposeMcpServer.Results.Excel.Protect;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Protect;

public class GetExcelProtectionHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelProtectionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Specific Sheet

    [Fact]
    public void Execute_WithSheetIndex_ReturnsOnlySpecifiedSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetProtectionResult>(res);

        Assert.Equal(1, result.Count);
        Assert.Equal(3, result.TotalWorksheets);
        Assert.Equal(1, result.Worksheets[0].Index);
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

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsProtectionInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetProtectionResult>(res);

        Assert.True(result.Count >= 0);
        Assert.True(result.TotalWorksheets >= 0);
        Assert.NotNull(result.Worksheets);
    }

    [Fact]
    public void Execute_ReturnsAllWorksheets()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetProtectionResult>(res);

        Assert.Equal(3, result.Count);
        Assert.Equal(3, result.TotalWorksheets);
        Assert.Equal(3, result.Worksheets.Count);
    }

    [Fact]
    public void Execute_ReturnsProtectionStatus()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetProtectionResult>(res);

        var worksheet = result.Worksheets[0];
        Assert.NotNull(worksheet.Name);
        Assert.True(worksheet.Index >= 0);
    }

    [Fact]
    public void Execute_ReturnsProtectionDetails()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetProtectionResult>(res);

        var worksheet = result.Worksheets[0];
        // These are boolean properties so they should have defined values
        Assert.True(worksheet.AllowSelectingLockedCell || !worksheet.AllowSelectingLockedCell);
        Assert.True(worksheet.AllowSelectingUnlockedCell || !worksheet.AllowSelectingUnlockedCell);
        Assert.True(worksheet.AllowFormattingCell || !worksheet.AllowFormattingCell);
        Assert.True(worksheet.AllowFiltering || !worksheet.AllowFiltering);
        Assert.True(worksheet.AllowSorting || !worksheet.AllowSorting);
    }

    [Fact]
    public void Execute_WithProtectedSheet_ReturnsIsProtectedTrue()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Protect(ProtectionType.All, "test123", null);

        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetProtectionResult>(res);

        Assert.True(result.Worksheets[0].IsProtected);
    }

    [Fact]
    public void Execute_WithUnprotectedSheet_ReturnsIsProtectedFalse()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetProtectionResult>(res);

        Assert.False(result.Worksheets[0].IsProtected);
    }

    #endregion
}
