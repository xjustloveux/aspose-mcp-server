using AsposeMcpServer.Handlers.Excel.Sheet;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Sheet;

public class CopyExcelSheetHandlerTests : ExcelHandlerTestBase
{
    private readonly CopyExcelSheetHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Copy()
    {
        Assert.Equal("copy", _handler.Operation);
    }

    #endregion

    #region Basic Copy Operations

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_CopiesSheetAtVariousIndices(int sheetIndex)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        var initialCount = workbook.Worksheets.Count;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", sheetIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount + 1, workbook.Worksheets.Count);
    }

    #endregion

    #region Preserve Original

    [Fact]
    public void Execute_PreservesOriginalSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Original Data";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Original Data", workbook.Worksheets[0].Cells["A1"].Value);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSheetNameInMessage()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Sheet1", result);
    }

    #endregion

    #region Target Index

    [Fact]
    public void Execute_WithTargetIndex_CopiesToPosition()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "targetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("1", result);
    }

    [Fact]
    public void Execute_WithoutTargetIndex_AppendsAtEnd()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(2, workbook.Worksheets.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sheetIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException(int invalidIndex)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidTargetIndex_ThrowsArgumentException(int invalidTarget)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "targetIndex", invalidTarget }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Copy To External File

    [Fact]
    public void Execute_WithCopyToPath_CreatesExternalFile()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test Data";
        var outputPath = CreateTestFilePath("copied.xlsx");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "copyToPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("external file", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithCopyToPath_PreservesSheetName()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Name = "CustomSheet";
        var outputPath = CreateTestFilePath("named_copy.xlsx");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "copyToPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("CustomSheet", result);
    }

    [Fact]
    public void Execute_WithCopyToPath_DoesNotModifyOriginal()
    {
        var workbook = CreateEmptyWorkbook();
        var initialCount = workbook.Worksheets.Count;
        var outputPath = CreateTestFilePath("external.xlsx");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "copyToPath", outputPath }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, workbook.Worksheets.Count);
    }

    #endregion
}
