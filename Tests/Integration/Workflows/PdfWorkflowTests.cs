using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Session;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;
using AsposeMcpServer.Tools.Session;

namespace AsposeMcpServer.Tests.Integration.Workflows;

/// <summary>
///     Integration tests for PDF document workflows.
/// </summary>
[Trait("Category", "Integration")]
public class PdfWorkflowTests : TestBase
{
    private readonly PdfAnnotationTool _annotationTool;
    private readonly PdfFileTool _fileTool;
    private readonly PdfPageTool _pageTool;
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;
    private readonly PdfTextTool _textTool;
    private readonly PdfWatermarkTool _watermarkTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfWorkflowTests" /> class.
    /// </summary>
    public PdfWorkflowTests()
    {
        var config = new SessionConfig { Enabled = true };
        _sessionManager = new DocumentSessionManager(config);
        var tempFileManager = new TempFileManager(config);
        _sessionTool = new DocumentSessionTool(_sessionManager, tempFileManager, new StdioSessionIdentityAccessor());
        _pageTool = new PdfPageTool(_sessionManager);
        _textTool = new PdfTextTool(_sessionManager);
        _watermarkTool = new PdfWatermarkTool(_sessionManager);
        _fileTool = new PdfFileTool(_sessionManager);
        _annotationTool = new PdfAnnotationTool(_sessionManager);
    }

    /// <summary>
    ///     Disposes of test resources.
    /// </summary>
    public override void Dispose()
    {
        _sessionManager.Dispose();
        base.Dispose();
    }

    #region Open-Edit-Save Workflow Tests

    /// <summary>
    ///     Verifies the complete open, edit, and save workflow for PDF documents.
    /// </summary>
    [Fact]
    public void Pdf_OpenEditSave_Workflow()
    {
        // Step 1: Create and open PDF
        var originalPath = CreatePdfDocument();
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Add a page
        _pageTool.Execute("add", sessionId: openData.SessionId);

        // Step 3: Save PDF
        var outputPath = CreateTestFilePath("pdf_workflow_output.pdf");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        // Step 4: Verify changes persisted
        var savedDoc = new Document(outputPath);
        Assert.True(savedDoc.Pages.Count >= 2);

        // Step 5: Close session
        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Watermark Workflow Tests

    /// <summary>
    ///     Verifies the workflow of adding a watermark to PDF.
    /// </summary>
    [Fact]
    public void Pdf_AddWatermark_Workflow()
    {
        // Step 1: Create and open PDF
        var originalPath = CreatePdfDocumentWithContent("Document content");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Add watermark
        _watermarkTool.Execute("add",
            sessionId: openData.SessionId,
            text: "CONFIDENTIAL",
            opacity: 0.3,
            fontSize: 48);

        // Step 3: Save and verify
        var outputPath = CreateTestFilePath("pdf_watermark_workflow.pdf");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Merge Workflow Tests

    /// <summary>
    ///     Verifies the workflow of merging multiple PDFs.
    /// </summary>
    [Fact]
    public void Pdf_MergeDocuments_Workflow()
    {
        // Step 1: Create multiple PDFs
        var pdf1Path = CreatePdfDocumentWithContent("First PDF content");
        var pdf2Path = CreatePdfDocumentWithContent("Second PDF content");
        var pdf3Path = CreatePdfDocumentWithContent("Third PDF content");

        // Step 2: Merge PDFs
        var outputPath = CreateTestFilePath("pdf_merged_output.pdf");
        _fileTool.Execute("merge",
            inputPaths: [pdf1Path, pdf2Path, pdf3Path],
            outputPath: outputPath);

        // Step 3: Verify merged PDF
        Assert.True(File.Exists(outputPath));

        var mergedDoc = new Document(outputPath);
        Assert.True(mergedDoc.Pages.Count >= 3);
    }

    #endregion

    #region Split Workflow Tests

    /// <summary>
    ///     Verifies the workflow of splitting a PDF.
    /// </summary>
    [Fact]
    public void Pdf_SplitPages_Workflow()
    {
        // Step 1: Create a multi-page PDF
        var originalPath = CreateMultiPagePdfDocument(3);

        // Step 2: Split PDF
        var outputDir = CreateTestFilePath("pdf_split_output");
        Directory.CreateDirectory(outputDir);

        _fileTool.Execute("split",
            originalPath,
            outputDir: outputDir,
            pagesPerFile: 1);

        // Step 3: Verify split files exist
        var splitFiles = Directory.GetFiles(outputDir, "*.pdf");
        Assert.True(splitFiles.Length >= 1);
    }

    #endregion

    #region Annotation Workflow Tests

    /// <summary>
    ///     Verifies the workflow of adding annotations to PDF.
    /// </summary>
    [Fact]
    public void Pdf_AddAnnotations_Workflow()
    {
        // Step 1: Create and open PDF
        var originalPath = CreatePdfDocumentWithContent("Document for annotation");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Add annotation
        _annotationTool.Execute("add",
            sessionId: openData.SessionId,
            pageIndex: 1,
            text: "This is an important note",
            x: 100,
            y: 700);

        // Step 3: Save and verify
        var outputPath = CreateTestFilePath("pdf_annotation_workflow.pdf");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));

        var savedDoc = new Document(outputPath);
        Assert.True(savedDoc.Pages[1].Annotations.Count > 0);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Page Management Workflow Tests

    /// <summary>
    ///     Verifies the workflow of adding and managing pages.
    /// </summary>
    [Fact]
    public void Pdf_AddPages_Workflow()
    {
        // Step 1: Create and open PDF
        var originalPath = CreatePdfDocument();
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Add multiple pages
        _pageTool.Execute("add", sessionId: openData.SessionId);
        _pageTool.Execute("add", sessionId: openData.SessionId);

        // Step 3: Save and verify
        var outputPath = CreateTestFilePath("pdf_pages_workflow.pdf");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        var savedDoc = new Document(outputPath);
        Assert.True(savedDoc.Pages.Count >= 3);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    /// <summary>
    ///     Verifies the workflow of extracting text from PDF.
    /// </summary>
    [Fact]
    public void Pdf_ExtractText_Workflow()
    {
        // Step 1: Create and open PDF with content
        var originalPath = CreatePdfDocumentWithContent("Sample PDF Content");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Extract text
        var extractResult = _textTool.Execute("extract", sessionId: openData.SessionId);

        // Step 3: Verify text was extracted
        Assert.NotNull(extractResult);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Helper Methods

    private string CreatePdfDocument()
    {
        var path = CreateTestFilePath($"pdf_{Guid.NewGuid()}.pdf");
        var doc = new Document();
        doc.Pages.Add();
        doc.Save(path);
        return path;
    }

    private string CreatePdfDocumentWithContent(string content)
    {
        var path = CreateTestFilePath($"pdf_content_{Guid.NewGuid()}.pdf");
        var doc = new Document();
        var page = doc.Pages.Add();
        page.Paragraphs.Add(new TextFragment(content));
        doc.Save(path);
        return path;
    }

    private string CreateMultiPagePdfDocument(int pageCount)
    {
        var path = CreateTestFilePath($"pdf_multipage_{Guid.NewGuid()}.pdf");
        var doc = new Document();
        for (var i = 1; i <= pageCount; i++)
        {
            var page = doc.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i} content"));
        }

        doc.Save(path);
        return path;
    }

    #endregion
}
