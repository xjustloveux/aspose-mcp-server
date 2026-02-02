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
[Collection("Workflow")]
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
        var config = new SessionConfig { Enabled = true, TempDirectory = Path.Combine(TestDir, "temp") };
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
        var originalPath = CreatePdfDocument();
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);
        Assert.False(string.IsNullOrEmpty(openData.SessionId));

        _pageTool.Execute("add", sessionId: openData.SessionId);

        var outputPath = CreateTestFilePath("pdf_workflow_output.pdf");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        var savedDoc = new Document(outputPath);
        Assert.True(savedDoc.Pages.Count >= 2);

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
        var originalPath = CreatePdfDocumentWithContent("Document content");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);
        Assert.False(string.IsNullOrEmpty(openData.SessionId));

        _watermarkTool.Execute("add",
            sessionId: openData.SessionId,
            text: "CONFIDENTIAL",
            opacity: 0.3,
            fontSize: 48);

        var outputPath = CreateTestFilePath("pdf_watermark_workflow.pdf");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        var savedDoc = new Document(outputPath);
        Assert.True(savedDoc.Pages.Count >= 1);
        Assert.True(new FileInfo(outputPath).Length > 0);

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
        var pdf1Path = CreatePdfDocumentWithContent("First PDF content");
        var pdf2Path = CreatePdfDocumentWithContent("Second PDF content");
        var pdf3Path = CreatePdfDocumentWithContent("Third PDF content");

        var outputPath = CreateTestFilePath("pdf_merged_output.pdf");
        _fileTool.Execute("merge",
            inputPaths: [pdf1Path, pdf2Path, pdf3Path],
            outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);

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
        var originalPath = CreateMultiPagePdfDocument(3);

        var outputDir = CreateTestFilePath("pdf_split_output");
        Directory.CreateDirectory(outputDir);

        _fileTool.Execute("split",
            originalPath,
            outputDir: outputDir,
            pagesPerFile: 1);

        var splitFiles = Directory.GetFiles(outputDir, "*.pdf");
        Assert.True(splitFiles.Length >= 1);
        foreach (var splitFile in splitFiles) Assert.True(new FileInfo(splitFile).Length > 0);
    }

    #endregion

    #region Annotation Workflow Tests

    /// <summary>
    ///     Verifies the workflow of adding annotations to PDF.
    /// </summary>
    [Fact]
    public void Pdf_AddAnnotations_Workflow()
    {
        var originalPath = CreatePdfDocumentWithContent("Document for annotation");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);
        Assert.False(string.IsNullOrEmpty(openData.SessionId));

        _annotationTool.Execute("add",
            sessionId: openData.SessionId,
            pageIndex: 1,
            text: "This is an important note",
            x: 100,
            y: 700);

        var outputPath = CreateTestFilePath("pdf_annotation_workflow.pdf");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);

        var savedDoc = new Document(outputPath);
        Assert.True(savedDoc.Pages.Count >= 1);
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
        var originalPath = CreatePdfDocument();
        var originalDoc = new Document(originalPath);
        var originalPageCount = originalDoc.Pages.Count;

        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);
        Assert.False(string.IsNullOrEmpty(openData.SessionId));

        _pageTool.Execute("add", sessionId: openData.SessionId);
        _pageTool.Execute("add", sessionId: openData.SessionId);

        var outputPath = CreateTestFilePath("pdf_pages_workflow.pdf");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        var savedDoc = new Document(outputPath);
        Assert.True(savedDoc.Pages.Count >= originalPageCount + 2);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    /// <summary>
    ///     Verifies the workflow of extracting text from PDF.
    /// </summary>
    [Fact]
    public void Pdf_ExtractText_Workflow()
    {
        var originalPath = CreatePdfDocumentWithContent("Sample PDF Content");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);
        Assert.False(string.IsNullOrEmpty(openData.SessionId));

        var extractResult = _textTool.Execute("extract", sessionId: openData.SessionId);

        Assert.NotNull(extractResult);
        Assert.True(extractResult.GetType().Name.StartsWith("FinalizedResult"),
            $"Expected FinalizedResult but got {extractResult.GetType().Name}");

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
