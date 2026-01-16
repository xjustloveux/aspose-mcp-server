using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tests.Helpers;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Unit tests for DocumentContext class
/// </summary>
public class DocumentContextTests : TestBase
{
    #region IsSession Property Tests

    [Fact]
    public void IsSession_WithFilePath_ShouldReturnFalse()
    {
        var docPath = CreateTestFilePath("test.docx");
        var doc = new Document();
        doc.Save(docPath);

        using var context = DocumentContext<Document>.Create(null, null, docPath);

        Assert.False(context.IsSession);
    }

    #endregion

    #region Save with Session Tests

    [Fact]
    public void Save_WithSession_ShouldMarkDirtyInsteadOfSaving()
    {
        var docPath = CreateTestFilePath("test_session_save.docx");
        var doc = new Document();
        doc.Save(docPath);

        var config = new SessionConfig { MaxSessions = 10 };
        var manager = new DocumentSessionManager(config);
        var sessionId = manager.OpenDocument(docPath);

        using var context = DocumentContext<Document>.Create(manager, sessionId, null);

        context.Save();

        var sessions = manager.ListSessions().ToList();
        Assert.Contains(sessions, s => s.SessionId == sessionId && s.IsDirty);

        manager.CloseDocument(sessionId, true);
    }

    #endregion

    #region GetOutputMessage Additional Tests

    [Fact]
    public void GetOutputMessage_WithOutputPathParameter_ShouldReturnOutputPath()
    {
        var docPath = CreateTestFilePath("test_msg.docx");
        var outputPath = CreateTestFilePath("output_msg.docx");
        var doc = new Document();
        doc.Save(docPath);

        using var context = DocumentContext<Document>.Create(null, null, docPath);

        var message = context.GetOutputMessage(outputPath);

        Assert.Contains("Output:", message);
        Assert.Contains(outputPath, message);
    }

    #endregion

    #region Create from File Path Tests

    [Fact]
    public void Create_WithWordDocument_ShouldLoadDocument()
    {
        var docPath = CreateTestFilePath("test.docx");
        var doc = new Document();
        doc.Save(docPath);

        using var context = DocumentContext<Document>.Create(null, null, docPath);

        Assert.NotNull(context.Document);
        Assert.False(context.IsSession);
        Assert.Equal(docPath, context.SourcePath);
    }

    [Fact]
    public void Create_WithExcelWorkbook_ShouldLoadDocument()
    {
        var xlsxPath = CreateTestFilePath("test.xlsx");
        using var workbook = new Workbook();
        workbook.Save(xlsxPath);

        using var context = DocumentContext<Workbook>.Create(null, null, xlsxPath);

        Assert.NotNull(context.Document);
        Assert.False(context.IsSession);
        Assert.Equal(xlsxPath, context.SourcePath);
    }

    [Fact]
    public void Create_WithPowerPointPresentation_ShouldLoadDocument()
    {
        var pptxPath = CreateTestFilePath("test.pptx");
        using var presentation = new Presentation();
        presentation.Save(pptxPath, SaveFormat.Pptx);

        using var context = DocumentContext<Presentation>.Create(null, null, pptxPath);

        Assert.NotNull(context.Document);
        Assert.False(context.IsSession);
        Assert.Equal(pptxPath, context.SourcePath);
    }

    [Fact]
    public void Create_WithPdfDocument_ShouldLoadDocument()
    {
        var pdfPath = CreateTestFilePath("test.pdf");
        using var pdfDoc = new Aspose.Pdf.Document();
        pdfDoc.Pages.Add();
        pdfDoc.Save(pdfPath);

        using var context = DocumentContext<Aspose.Pdf.Document>.Create(null, null, pdfPath);

        Assert.NotNull(context.Document);
        Assert.False(context.IsSession);
        Assert.Equal(pdfPath, context.SourcePath);
    }

    [Fact]
    public void Create_WithNullPathAndNoSession_ShouldThrow()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            DocumentContext<Document>.Create(null, null, null));

        Assert.Contains("must be provided", ex.Message);
    }

    [Fact]
    public void Create_WithEmptyPathAndNoSession_ShouldThrow()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            DocumentContext<Document>.Create(null, null, ""));

        Assert.Contains("must be provided", ex.Message);
    }

    #endregion

    #region Create from Session Tests

    [Fact]
    public void Create_WithSessionId_ButNoManager_ShouldThrow()
    {
        var ex = Assert.Throws<InvalidOperationException>(() =>
            DocumentContext<Document>.Create(null, "sess_123", null));

        Assert.Contains("Session management is not enabled", ex.Message);
    }

    [Fact]
    public void Create_WithSessionId_ShouldGetDocumentFromManager()
    {
        var docPath = CreateTestFilePath("test.docx");
        var doc = new Document();
        doc.Save(docPath);

        var config = new SessionConfig
        {
            MaxSessions = 10
        };
        var manager = new DocumentSessionManager(config);
        var sessionId = manager.OpenDocument(docPath);

        using var context = DocumentContext<Document>.Create(manager, sessionId, null);

        Assert.NotNull(context.Document);
        Assert.True(context.IsSession);
        Assert.Null(context.SourcePath);

        manager.CloseDocument(sessionId);
    }

    #endregion

    #region Save Tests

    [Fact]
    public void Save_WithFilePath_ShouldSaveToPath()
    {
        var docPath = CreateTestFilePath("test.docx");
        var outputPath = CreateTestFilePath("output.docx");
        var doc = new Document();
        doc.Save(docPath);

        using var context = DocumentContext<Document>.Create(null, null, docPath);
        var builder = new DocumentBuilder(context.Document);
        builder.Write("Modified content");

        context.Save(outputPath);

        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Save_WithoutOutputPath_ShouldSaveToSourcePath()
    {
        var docPath = CreateTestFilePath("test.docx");
        var doc = new Document();
        doc.Save(docPath);

        using var context = DocumentContext<Document>.Create(null, null, docPath);
        var builder = new DocumentBuilder(context.Document);
        builder.Write("Modified content");

        context.Save();

        Assert.True(File.Exists(docPath));
    }

    [Fact]
    public void Save_Excel_ShouldSaveToPath()
    {
        var xlsxPath = CreateTestFilePath("test.xlsx");
        var outputPath = CreateTestFilePath("output.xlsx");
        using var workbook = new Workbook();
        workbook.Save(xlsxPath);

        using var context = DocumentContext<Workbook>.Create(null, null, xlsxPath);
        context.Document.Worksheets[0].Cells["A1"].PutValue("Test");

        context.Save(outputPath);

        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Save_PowerPoint_ShouldSaveToPath()
    {
        var pptxPath = CreateTestFilePath("test.pptx");
        var outputPath = CreateTestFilePath("output.pptx");
        using var presentation = new Presentation();
        presentation.Save(pptxPath, SaveFormat.Pptx);

        using var context = DocumentContext<Presentation>.Create(null, null, pptxPath);
        context.Document.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        context.Save(outputPath);

        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Save_Pdf_ShouldSaveToPath()
    {
        var pdfPath = CreateTestFilePath("test.pdf");
        var outputPath = CreateTestFilePath("output.pdf");
        using var pdfDoc = new Aspose.Pdf.Document();
        pdfDoc.Pages.Add();
        pdfDoc.Save(pdfPath);

        using var context = DocumentContext<Aspose.Pdf.Document>.Create(null, null, pdfPath);

        context.Save(outputPath);

        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region MarkDirty Tests

    [Fact]
    public void MarkDirty_WithSession_ShouldMarkSessionDirty()
    {
        var docPath = CreateTestFilePath("test.docx");
        var doc = new Document();
        doc.Save(docPath);

        var config = new SessionConfig
        {
            MaxSessions = 10
        };
        var manager = new DocumentSessionManager(config);
        var sessionId = manager.OpenDocument(docPath);

        using var context = DocumentContext<Document>.Create(manager, sessionId, null);
        context.MarkDirty();

        var sessions = manager.ListSessions();
        Assert.Contains(sessions, s => s.SessionId == sessionId && s.IsDirty);

        manager.CloseDocument(sessionId, true);
    }

    [Fact]
    public void MarkDirty_WithFilePath_ShouldNotThrow()
    {
        var docPath = CreateTestFilePath("test.docx");
        var doc = new Document();
        doc.Save(docPath);

        using var context = DocumentContext<Document>.Create(null, null, docPath);

        var exception = Record.Exception(() => context.MarkDirty());

        Assert.Null(exception);
    }

    #endregion

    #region GetOutputMessage Tests

    [Fact]
    public void GetOutputMessage_WithFilePath_ShouldReturnPathMessage()
    {
        var docPath = CreateTestFilePath("test.docx");
        var doc = new Document();
        doc.Save(docPath);

        using var context = DocumentContext<Document>.Create(null, null, docPath);

        var message = context.GetOutputMessage();

        Assert.Contains("Output:", message);
        Assert.Contains(docPath, message);
    }

    [Fact]
    public void GetOutputMessage_WithSession_ShouldReturnSessionMessage()
    {
        var docPath = CreateTestFilePath("test.docx");
        var doc = new Document();
        doc.Save(docPath);

        var config = new SessionConfig
        {
            MaxSessions = 10
        };
        var manager = new DocumentSessionManager(config);
        var sessionId = manager.OpenDocument(docPath);

        using var context = DocumentContext<Document>.Create(manager, sessionId, null);

        var message = context.GetOutputMessage();

        Assert.Contains("session", message);
        Assert.Contains(sessionId, message);

        manager.CloseDocument(sessionId, true);
    }

    #endregion

    #region Dispose Tests

    [Fact]
    public void Dispose_WithOwnedDocument_ShouldDisposeDocument()
    {
        var docPath = CreateTestFilePath("test.docx");
        var doc = new Document();
        doc.Save(docPath);

        var context = DocumentContext<Document>.Create(null, null, docPath);
        _ = context.Document;

        context.Dispose();

        Assert.True(true);
    }

    [Fact]
    public void Dispose_MultipleTimes_ShouldNotThrow()
    {
        var docPath = CreateTestFilePath("test.docx");
        var doc = new Document();
        doc.Save(docPath);

        var context = DocumentContext<Document>.Create(null, null, docPath);

        context.Dispose();
        var exception = Record.Exception(() => context.Dispose());

        Assert.Null(exception);
    }

    #endregion

    #region Load with Password Tests

    [Fact]
    public void Create_WordDocument_WithWrongPassword_ShouldLoadWithoutPassword()
    {
        var docPath = CreateTestFilePath("test_word_pwd.docx");
        var doc = new Document();
        doc.Save(docPath);

        using var context = DocumentContext<Document>.Create(null, null, docPath, password: "wrongpassword");

        Assert.NotNull(context.Document);
    }

    [Fact]
    public void Create_ExcelWorkbook_WithWrongPassword_ShouldLoadWithoutPassword()
    {
        var xlsxPath = CreateTestFilePath("test_excel_pwd.xlsx");
        using var workbook = new Workbook();
        workbook.Save(xlsxPath);

        using var context = DocumentContext<Workbook>.Create(null, null, xlsxPath, password: "wrongpassword");

        Assert.NotNull(context.Document);
    }

    [Fact]
    public void Create_PowerPointPresentation_WithWrongPassword_ShouldLoadWithoutPassword()
    {
        var pptxPath = CreateTestFilePath("test_ppt_pwd.pptx");
        using var presentation = new Presentation();
        presentation.Save(pptxPath, SaveFormat.Pptx);

        using var context = DocumentContext<Presentation>.Create(null, null, pptxPath, password: "wrongpassword");

        Assert.NotNull(context.Document);
    }

    [Fact]
    public void Create_PdfDocument_WithWrongPassword_ShouldLoadWithoutPassword()
    {
        var pdfPath = CreateTestFilePath("test_pdf_pwd.pdf");
        using var pdfDoc = new Aspose.Pdf.Document();
        pdfDoc.Pages.Add();
        pdfDoc.Save(pdfPath);

        using var context = DocumentContext<Aspose.Pdf.Document>.Create(null, null, pdfPath, password: "wrongpassword");

        Assert.NotNull(context.Document);
    }

    #endregion

    #region Create with Identity Accessor Tests

    [Fact]
    public void Create_WithIdentityAccessor_ShouldUseIdentity()
    {
        var docPath = CreateTestFilePath("test_identity.docx");
        var doc = new Document();
        doc.Save(docPath);

        var config = new SessionConfig
        {
            MaxSessions = 10,
            IsolationMode = SessionIsolationMode.Group
        };
        var manager = new DocumentSessionManager(config);
        var identity = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var sessionId = manager.OpenDocument(docPath, identity);

        var mockAccessor = new TestIdentityAccessor(identity);
        using var context = DocumentContext<Document>.Create(manager, sessionId, null, mockAccessor);

        Assert.NotNull(context.Document);

        manager.CloseDocument(sessionId, identity, true);
    }

    private class TestIdentityAccessor : ISessionIdentityAccessor
    {
        private readonly SessionIdentity _identity;

        public TestIdentityAccessor(SessionIdentity identity)
        {
            _identity = identity;
        }

        public SessionIdentity GetCurrentIdentity()
        {
            return _identity;
        }
    }

    #endregion
}
