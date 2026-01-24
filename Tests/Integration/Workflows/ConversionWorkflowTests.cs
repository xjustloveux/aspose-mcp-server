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
public class ConversionWorkflowTests : TestBase
{
    private readonly ConvertDocumentTool _convertDocumentTool;
    private readonly ConvertToPdfTool _convertToPdfTool;
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ConversionWorkflowTests" /> class.
    /// </summary>
    public ConversionWorkflowTests()
    {
        var config = new SessionConfig { Enabled = true };
        _sessionManager = new DocumentSessionManager(config);
        var tempFileManager = new TempFileManager(config);
        _sessionTool = new DocumentSessionTool(_sessionManager, tempFileManager, new StdioSessionIdentityAccessor());
        _convertToPdfTool = new ConvertToPdfTool(_sessionManager);
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
        // Step 1: Create Excel document
        var excelPath = CreateExcelDocument();
        var pdfPath = CreateTestFilePath("excel_to_pdf_output.pdf");

        // Step 2: Convert to PDF
        var result = _convertToPdfTool.Execute(excelPath, outputPath: pdfPath);

        // Step 3: Verify PDF was created
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
        // Step 1: Create PowerPoint document
        var pptPath = CreatePowerPointDocument();
        var pdfPath = CreateTestFilePath("ppt_to_pdf_output.pdf");

        // Step 2: Convert to PDF
        var result = _convertToPdfTool.Execute(pptPath, outputPath: pdfPath);

        // Step 3: Verify PDF was created
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
        // Step 1: Create multiple documents
        var wordPath = CreateWordDocument("Word content for batch");
        var excelPath = CreateExcelDocument();
        var pptPath = CreatePowerPointDocument();

        var inputFiles = new[] { wordPath, excelPath, pptPath };
        var outputFiles = new List<string>();

        // Step 2: Convert each file to PDF
        foreach (var inputFile in inputFiles)
        {
            var outputPath = CreateTestFilePath($"batch_{Path.GetFileNameWithoutExtension(inputFile)}.pdf");
            var result = _convertToPdfTool.Execute(inputFile, outputPath: outputPath);
            outputFiles.Add(result.OutputPath);
        }

        // Step 3: Verify all PDFs were created
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
        // Step 1: Create Word document
        var wordPath = CreateWordDocument("Content for PDF conversion");
        var pdfPath = CreateTestFilePath("word_to_pdf_output.pdf");

        // Step 2: Convert to PDF
        var result = _convertToPdfTool.Execute(wordPath, outputPath: pdfPath);

        // Step 3: Verify PDF was created
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
        // Step 1: Create and open Word document
        var wordPath = CreateWordDocument("Session content for PDF");
        var openResult = _sessionTool.Execute("open", wordPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Convert to PDF using session
        var pdfPath = CreateTestFilePath("session_to_pdf_output.pdf");
        var result = _convertToPdfTool.Execute(sessionId: openData.SessionId, outputPath: pdfPath);

        // Step 3: Verify PDF was created
        Assert.True(File.Exists(pdfPath));
        Assert.Equal(pdfPath, result.OutputPath);

        // Step 4: Close session
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
        // Step 1: Create Word document
        var wordPath = CreateWordDocument("Content for HTML conversion");
        var htmlPath = CreateTestFilePath("word_to_html_output.html");

        // Step 2: Convert to HTML
        var result = _convertDocumentTool.Execute(wordPath, outputPath: htmlPath);

        // Step 3: Verify HTML was created
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
        // Step 1: Create Excel document
        var excelPath = CreateExcelDocument();
        var htmlPath = CreateTestFilePath("excel_to_html_output.html");

        // Step 2: Convert to HTML
        var result = _convertDocumentTool.Execute(excelPath, outputPath: htmlPath);

        // Step 3: Verify HTML was created
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
