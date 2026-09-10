using System.Text.Json;
using AsposeMcpServer.Handlers.Excel.Render;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Render;

/// <summary>
///     Found by the container smoke test for DEPLOY-03. Rendering a worksheet with no printable
///     content produced <c>PageCount = 0</c>, wrote nothing, and still reported success with the
///     requested path in <c>OutputPaths</c> — so the caller was handed a path to a file that does
///     not exist. This is the defect RB-27 fixed inside <c>DocumentConverter</c>; this handler
///     builds its own list and was missed.
/// </summary>
public class RenderSheetEmptyWorksheetTests : ExcelHandlerTestBase
{
    private readonly RenderSheetExcelHandler _handler = new();

    [Fact]
    public void Execute_WithAnEmptyWorksheet_ShouldRefuseInsteadOfReportingAPhantomFile()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var outputPath = CreateTestFilePath("empty_sheet.png");
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "outputPath", outputPath }
        });

        var exception = Assert.ThrowsAny<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Contains("no printable content", exception.Message);
        Assert.False(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithContent_ShouldStillRenderAndReportTheFileItWrote()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("visible content");
        var context = CreateContext(workbook);
        var outputPath = CreateTestFilePath("filled_sheet.png");
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonSerializer.Serialize(result);
        Assert.Contains(Path.GetFileName(outputPath), json);
        Assert.True(File.Exists(outputPath));
    }
}
