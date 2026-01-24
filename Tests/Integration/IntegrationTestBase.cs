using Aspose.Cells;
using Aspose.Pdf.Text;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Tests.Infrastructure;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Integration;

/// <summary>
///     Base class for integration tests providing document creation and session management capabilities.
/// </summary>
public abstract class IntegrationTestBase : TestBase
{
    /// <summary>
    ///     Creates a test Word document and returns its path.
    /// </summary>
    /// <param name="content">The text content for the document.</param>
    /// <param name="fileName">Optional custom file name.</param>
    /// <returns>The full path to the created file.</returns>
    protected string CreateWordDocument(string content = "Test Content", string? fileName = null)
    {
        var path = CreateTestFilePath(fileName ?? $"word_{Guid.NewGuid()}.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(path);
        return path;
    }

    /// <summary>
    ///     Creates a test Excel document and returns its path.
    /// </summary>
    /// <param name="fileName">Optional custom file name.</param>
    /// <returns>The full path to the created file.</returns>
    protected string CreateExcelDocument(string? fileName = null)
    {
        var path = CreateTestFilePath(fileName ?? $"excel_{Guid.NewGuid()}.xlsx");
        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(path);
        return path;
    }

    /// <summary>
    ///     Creates a test Excel document with data and returns its path.
    /// </summary>
    /// <param name="fileName">Optional custom file name.</param>
    /// <returns>The full path to the created file.</returns>
    protected string CreateExcelDocumentWithData(string? fileName = null)
    {
        var path = CreateTestFilePath(fileName ?? $"excel_data_{Guid.NewGuid()}.xlsx");
        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["B1"].Value = 20;
        workbook.Save(path);
        return path;
    }

    /// <summary>
    ///     Creates a test PowerPoint document and returns its path.
    /// </summary>
    /// <param name="fileName">Optional custom file name.</param>
    /// <returns>The full path to the created file.</returns>
    protected string CreatePowerPointDocument(string? fileName = null)
    {
        var path = CreateTestFilePath(fileName ?? $"ppt_{Guid.NewGuid()}.pptx");
        using var pres = new Presentation();
        pres.Save(path, SaveFormat.Pptx);
        return path;
    }

    /// <summary>
    ///     Creates a test PDF document and returns its path.
    /// </summary>
    /// <param name="fileName">Optional custom file name.</param>
    /// <returns>The full path to the created file.</returns>
    protected string CreatePdfDocument(string? fileName = null)
    {
        var path = CreateTestFilePath(fileName ?? $"pdf_{Guid.NewGuid()}.pdf");
        var doc = new Aspose.Pdf.Document();
        var page = doc.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test Content"));
        doc.Save(path);
        return path;
    }

    /// <summary>
    ///     Reads the content of a Word document.
    /// </summary>
    /// <param name="path">The path to the Word document.</param>
    /// <returns>The text content of the document.</returns>
    protected static string ReadWordDocumentContent(string path)
    {
        var doc = new Document(path);
        return doc.GetText();
    }

    /// <summary>
    ///     Reads a cell value from an Excel document.
    /// </summary>
    /// <param name="path">The path to the Excel document.</param>
    /// <param name="cell">The cell reference (e.g., "A1").</param>
    /// <returns>The cell value as a string.</returns>
    protected static string ReadExcelCellValue(string path, string cell)
    {
        using var workbook = new Workbook(path);
        return workbook.Worksheets[0].Cells[cell].StringValue;
    }
}
