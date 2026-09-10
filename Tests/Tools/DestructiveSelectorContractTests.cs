using Aspose.Cells;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;
using AsposeMcpServer.Tools.Pdf;
using PdfDocument = Aspose.Pdf.Document;

namespace AsposeMcpServer.Tests.Tools;

/// <summary>
///     Guards NEW-CONTRACT-01. Several destructive operations described their selector as required
///     but declared it as a non-nullable parameter defaulting to 0, and the tool always forwarded
///     that value. Omitting the selector was therefore indistinguishable from asking for item 0, so
///     the handler's own "selector is required" check could never fire and the call silently
///     removed the first item instead of refusing.
/// </summary>
public class DestructiveSelectorContractTests : TestBase
{
    /// <summary>Creates a workbook with the requested number of sheets.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <param name="sheetCount">Total sheets to create.</param>
    /// <returns>The path written.</returns>
    private string CreateWorkbook(string fileName, int sheetCount)
    {
        var path = CreateTestFilePath(fileName);
        using var workbook = new Workbook();
        while (workbook.Worksheets.Count < sheetCount) workbook.Worksheets.Add();
        for (var i = 0; i < workbook.Worksheets.Count; i++)
            workbook.Worksheets[i].Cells["A1"].PutValue($"sheet{i}");
        workbook.Save(path);
        return path;
    }

    /// <summary>Creates a PDF with the requested number of pages.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <param name="pageCount">Total pages to create.</param>
    /// <returns>The path written.</returns>
    private string CreatePdf(string fileName, int pageCount)
    {
        var path = CreateTestFilePath(fileName);
        using var pdf = new PdfDocument();
        for (var i = 0; i < pageCount; i++) pdf.Pages.Add();
        pdf.Save(path);
        return path;
    }

    [Fact]
    public void ExcelSheetDelete_WithoutIndex_ShouldBeRefused()
    {
        var tool = new ExcelSheetTool();
        var path = CreateWorkbook("sheets.xlsx", 3);

        Assert.ThrowsAny<ArgumentException>(() => tool.Execute("delete", path));

        using var after = new Workbook(path);
        var labels = after.Worksheets
            .Select(sheet => sheet.Cells["A1"].StringValue)
            .ToList();
        Assert.Contains("sheet0", labels);
        Assert.Contains("sheet1", labels);
        Assert.Contains("sheet2", labels);
    }

    [Fact]
    public void ExcelSheetDelete_WithIndex_ShouldStillWork()
    {
        var tool = new ExcelSheetTool();
        var path = CreateWorkbook("sheets_ok.xlsx", 3);
        var outputPath = CreateTestFilePath("sheets_ok_out.xlsx");

        tool.Execute("delete", path, outputPath: outputPath, sheetIndex: 1);

        using var after = new Workbook(outputPath);
        var labels = after.Worksheets
            .Select(sheet => sheet.Cells["A1"].StringValue)
            .ToList();
        Assert.DoesNotContain("sheet1", labels);
        Assert.Contains("sheet0", labels);
        Assert.Contains("sheet2", labels);
    }

    [Fact]
    public void ExcelRowDelete_WithoutIndex_ShouldBeRefused()
    {
        var tool = new ExcelRowColumnTool();
        var path = CreateWorkbook("rows.xlsx", 1);

        Assert.ThrowsAny<ArgumentException>(() => tool.Execute("delete_row", path));
    }

    [Fact]
    public void ExcelColumnDelete_WithoutIndex_ShouldBeRefused()
    {
        var tool = new ExcelRowColumnTool();
        var path = CreateWorkbook("cols.xlsx", 1);

        Assert.ThrowsAny<ArgumentException>(() => tool.Execute("delete_column", path));
    }

    [Fact]
    public void ExcelChartDelete_WithoutIndex_ShouldBeRefused()
    {
        var tool = new ExcelChartTool();
        var path = CreateWorkbook("charts.xlsx", 1);

        Assert.ThrowsAny<ArgumentException>(() => tool.Execute("delete", path));
    }

    [Fact]
    public void PdfSignatureDelete_WithoutSelector_ShouldBeRefused()
    {
        var tool = new PdfSignatureTool();
        var path = CreatePdf("signature.pdf", 1);

        Assert.ThrowsAny<ArgumentException>(() => tool.Execute("delete", path));
    }

    [Fact]
    public void PdfLinkDelete_WithoutIndex_ShouldBeRefused()
    {
        var tool = new PdfLinkTool();
        var path = CreatePdf("links.pdf", 1);

        Assert.ThrowsAny<ArgumentException>(() => tool.Execute("delete", path, pageIndex: 1));
    }
}
