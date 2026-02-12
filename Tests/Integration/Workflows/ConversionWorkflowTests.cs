using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Session;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Conversion;
using AsposeMcpServer.Tools.Session;
using Document = Aspose.Pdf.Document;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Integration.Workflows;

/// <summary>
///     Integration tests for document conversion workflows.
/// </summary>
[Trait("Category", "Integration")]
[Collection("Workflow")]
public class ConversionWorkflowTests : TestBase
{
    private readonly ConvertDocumentTool _convertDocumentTool;
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ConversionWorkflowTests" /> class.
    /// </summary>
    public ConversionWorkflowTests()
    {
        var config = new SessionConfig { Enabled = true, TempDirectory = Path.Combine(TestDir, "temp") };
        _sessionManager = new DocumentSessionManager(config);
        var tempFileManager = new TempFileManager(config);
        _sessionTool = new DocumentSessionTool(_sessionManager, tempFileManager, new StdioSessionIdentityAccessor());
        _convertDocumentTool = new ConvertDocumentTool(_sessionManager);
    }

    /// <summary>
    ///     Disposes of test resources.
    /// </summary>
    public override void Dispose()
    {
        _sessionManager.Dispose();
        base.Dispose();
    }

    #region Excel to PDF Conversion Tests

    /// <summary>
    ///     Verifies Excel to PDF conversion workflow.
    /// </summary>
    [Fact]
    public void Conversion_ExcelToPdf_Workflow()
    {
        var excelPath = CreateExcelDocument();
        var pdfPath = CreateTestFilePath("excel_to_pdf_output.pdf");

        var result = _convertDocumentTool.Execute(excelPath, outputPath: pdfPath);

        Assert.True(File.Exists(pdfPath));
        Assert.Equal(pdfPath, result.OutputPath);
    }

    #endregion

    #region PowerPoint to PDF Conversion Tests

    /// <summary>
    ///     Verifies PowerPoint to PDF conversion workflow.
    /// </summary>
    [Fact]
    public void Conversion_PowerPointToPdf_Workflow()
    {
        var pptPath = CreatePowerPointDocument();
        var pdfPath = CreateTestFilePath("ppt_to_pdf_output.pdf");

        var result = _convertDocumentTool.Execute(pptPath, outputPath: pdfPath);

        Assert.True(File.Exists(pdfPath));
        Assert.Equal(pdfPath, result.OutputPath);
    }

    #endregion

    #region Batch Conversion Tests

    /// <summary>
    ///     Verifies batch conversion of multiple files to PDF.
    /// </summary>
    [Fact]
    public void Conversion_BatchConvert_Workflow()
    {
        var wordPath = CreateWordDocument("Word content for batch");
        var excelPath = CreateExcelDocument();
        var pptPath = CreatePowerPointDocument();

        var inputFiles = new[] { wordPath, excelPath, pptPath };
        var outputFiles = new List<string>();

        foreach (var inputFile in inputFiles)
        {
            var outputPath = CreateTestFilePath($"batch_{Path.GetFileNameWithoutExtension(inputFile)}.pdf");
            var result = _convertDocumentTool.Execute(inputFile, outputPath: outputPath);
            outputFiles.Add(result.OutputPath);
        }

        Assert.Equal(3, outputFiles.Count);
        foreach (var outputFile in outputFiles) Assert.True(File.Exists(outputFile));
    }

    #endregion

    #region Word to PDF Conversion Tests

    /// <summary>
    ///     Verifies Word to PDF conversion workflow using file path.
    /// </summary>
    [Fact]
    public void Conversion_WordToPdf_FromPath_Workflow()
    {
        var wordPath = CreateWordDocument("Content for PDF conversion");
        var pdfPath = CreateTestFilePath("word_to_pdf_output.pdf");

        var result = _convertDocumentTool.Execute(wordPath, outputPath: pdfPath);

        Assert.True(File.Exists(pdfPath));
        Assert.Equal(pdfPath, result.OutputPath);

        var pdfDoc = new Document(pdfPath);
        Assert.True(pdfDoc.Pages.Count >= 1);
    }

    /// <summary>
    ///     Verifies Word to PDF conversion workflow using session.
    /// </summary>
    [Fact]
    public void Conversion_WordToPdf_FromSession_Workflow()
    {
        var wordPath = CreateWordDocument("Session content for PDF");
        var openResult = _sessionTool.Execute("open", wordPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        var pdfPath = CreateTestFilePath("session_to_pdf_output.pdf");
        var result = _convertDocumentTool.Execute(sessionId: openData.SessionId, outputPath: pdfPath);

        Assert.True(File.Exists(pdfPath));
        Assert.Equal(pdfPath, result.OutputPath);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region HTML Conversion Tests

    /// <summary>
    ///     Verifies Word to HTML conversion workflow.
    /// </summary>
    [Fact]
    public void Conversion_WordToHtml_Workflow()
    {
        var wordPath = CreateWordDocument("Content for HTML conversion");
        var htmlPath = CreateTestFilePath("word_to_html_output.html");

        var result = _convertDocumentTool.Execute(wordPath, outputPath: htmlPath);

        Assert.True(File.Exists(htmlPath));
        Assert.Equal(htmlPath, result.OutputPath);

        var htmlContent = File.ReadAllText(htmlPath);
        Assert.Contains("html", htmlContent.ToLower());
    }

    /// <summary>
    ///     Verifies Excel to HTML conversion workflow.
    /// </summary>
    [Fact]
    public void Conversion_ExcelToHtml_Workflow()
    {
        var excelPath = CreateExcelDocument();
        var htmlPath = CreateTestFilePath("excel_to_html_output.html");

        var result = _convertDocumentTool.Execute(excelPath, outputPath: htmlPath);

        Assert.True(File.Exists(htmlPath));
        Assert.Equal(htmlPath, result.OutputPath);
    }

    #endregion

    #region Helper Methods

    private string CreateWordDocument(string content = "Test Content")
    {
        var path = CreateTestFilePath($"word_{Guid.NewGuid()}.docx");
        var doc = new Aspose.Words.Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(path);
        return path;
    }

    private string CreateExcelDocument()
    {
        var path = CreateTestFilePath($"excel_{Guid.NewGuid()}.xlsx");
        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test Data";
        workbook.Save(path);
        return path;
    }

    private string CreatePowerPointDocument()
    {
        var path = CreateTestFilePath($"ppt_{Guid.NewGuid()}.pptx");
        using var pres = new Presentation();
        pres.Save(path, SaveFormat.Pptx);
        return path;
    }

    #endregion
}
