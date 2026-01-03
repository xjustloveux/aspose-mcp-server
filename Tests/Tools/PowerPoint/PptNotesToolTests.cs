using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptNotesToolTests : TestBase
{
    private readonly PptNotesTool _tool;

    public PptNotesToolTests()
    {
        _tool = new PptNotesTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName, int slideCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithNotes(string fileName, string notesText)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var notesSlide = presentation.Slides[0].NotesSlideManager.AddNotesSlide();
        if (notesSlide.NotesTextFrame != null)
            notesSlide.NotesTextFrame.Text = notesText;
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void Set_ShouldSetNotes()
    {
        var pptPath = CreateTestPresentation("test_set_notes.pptx");
        var outputPath = CreateTestFilePath("test_set_notes_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, notes: "Speaker notes for this slide",
            outputPath: outputPath);
        Assert.Contains("Notes set", result);
        using var presentation = new Presentation(outputPath);
        var notesSlide = presentation.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
            // Evaluation mode: content may be truncated with watermark
            Assert.False(string.IsNullOrEmpty(notesSlide.NotesTextFrame.Text));
        else
            // Licensed mode: strict content verification
            Assert.Contains("Speaker notes", notesSlide.NotesTextFrame.Text);
    }

    [Fact]
    public void Set_ShouldReplaceExistingNotes()
    {
        var pptPath = CreatePresentationWithNotes("test_set_replace.pptx", "Original notes");
        var outputPath = CreateTestFilePath("test_set_replace_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, notes: "Updated notes content",
            outputPath: outputPath);
        Assert.Contains("Notes set", result);
        using var presentation = new Presentation(outputPath);
        var notesSlide = presentation.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
            // Evaluation mode: just verify notes were changed (not equal to original)
            Assert.DoesNotContain("Original notes", notesSlide.NotesTextFrame.Text);
        else
            // Licensed mode: strict content verification
            Assert.Equal("Updated notes content", notesSlide.NotesTextFrame.Text);
    }

    [Fact]
    public void Get_WithSlideIndex_ShouldReturnSingleSlideNotes()
    {
        var pptPath = CreatePresentationWithNotes("test_get_notes.pptx", "Test notes content");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("slideIndex").GetInt32());
        Assert.True(json.RootElement.GetProperty("hasNotes").GetBoolean());

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
            // Evaluation mode: content may be truncated with watermark
            Assert.False(string.IsNullOrEmpty(json.RootElement.GetProperty("notes").GetString()));
        else
            // Licensed mode: strict content verification
            Assert.Contains("Test notes", json.RootElement.GetProperty("notes").GetString());
    }

    [Fact]
    public void Get_WithoutSlideIndex_ShouldReturnAllNotes()
    {
        var pptPath = CreateTestPresentation("test_get_all_notes.pptx", 3);
        var result = _tool.Execute("get", pptPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(3, json.RootElement.GetProperty("slides").GetArrayLength());
    }

    [Fact]
    public void Clear_WithSlideIndices_ShouldClearSpecificSlides()
    {
        var pptPath = CreatePresentationWithNotes("test_clear_notes.pptx", "Notes to clear");
        var outputPath = CreateTestFilePath("test_clear_notes_output.pptx");
        var result = _tool.Execute("clear", pptPath, slideIndices: [0], outputPath: outputPath);
        Assert.Contains("Cleared", result);
        using var presentation = new Presentation(outputPath);
        var notesSlide = presentation.Slides[0].NotesSlideManager.NotesSlide;
        Assert.True(string.IsNullOrEmpty(notesSlide?.NotesTextFrame?.Text));
    }

    [Fact]
    public void Clear_WithoutSlideIndices_ShouldClearAllSlides()
    {
        var pptPath = CreateTestPresentation("test_clear_all.pptx", 3);
        var outputPath = CreateTestFilePath("test_clear_all_output.pptx");
        var result = _tool.Execute("clear", pptPath, outputPath: outputPath);
        Assert.Contains("3 targeted", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetHeaderFooter_ShouldSetNotesHeaderFooter()
    {
        var pptPath = CreateTestPresentation("test_notes_header_footer.pptx");
        var outputPath = CreateTestFilePath("test_notes_header_footer_output.pptx");
        var result = _tool.Execute("set_header_footer", pptPath, headerText: "Notes Header", footerText: "Notes Footer",
            outputPath: outputPath);
        Assert.Contains("header", result);
        Assert.Contains("footer", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetHeaderFooter_WithAllParameters_ShouldSetAll()
    {
        var pptPath = CreateTestPresentation("test_notes_hf_all.pptx");
        var outputPath = CreateTestFilePath("test_notes_hf_all_output.pptx");
        var result = _tool.Execute("set_header_footer", pptPath, headerText: "Header", footerText: "Footer",
            dateText: "2024-12-28", showPageNumber: true, outputPath: outputPath);
        Assert.Contains("header", result);
        Assert.Contains("footer", result);
        Assert.Contains("date", result);
        Assert.Contains("page number shown", result);
    }

    [Fact]
    public void SetHeaderFooter_HidePageNumber_ShouldHide()
    {
        var pptPath = CreateTestPresentation("test_notes_hide_page.pptx");
        var outputPath = CreateTestFilePath("test_notes_hide_page_output.pptx");
        var result = _tool.Execute("set_header_footer", pptPath, showPageNumber: false, outputPath: outputPath);
        Assert.Contains("page number hidden", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Set_WithInvalidSlideIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_set_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set", pptPath, slideIndex: 99, notes: "Test notes"));
        Assert.Contains("slide", ex.Message.ToLower());
    }

    [Fact]
    public void Get_WithInvalidSlideIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_get_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get", pptPath, slideIndex: 99));
        Assert.Contains("slide", ex.Message.ToLower());
    }

    [Fact]
    public void Clear_WithInvalidSlideIndices_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_clear_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("clear", pptPath, slideIndices: [0, 99]));
        Assert.Contains("Invalid slide indices", ex.Message);
    }

    [Fact]
    public void SetHeaderFooter_OnEmptyPresentation_ShouldThrow()
    {
        // Arrange - Create empty presentation
        var pptPath = CreateTestFilePath("test_notes_empty.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Slides.RemoveAt(0);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath, headerText: "Header"));
        Assert.Contains("no slides", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void Get_WithSessionId_ShouldReturnNotes()
    {
        var pptPath = CreatePresentationWithNotes("test_session_get_notes.pptx", "Session test notes");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("slideIndex").GetInt32());
        Assert.True(json.RootElement.GetProperty("hasNotes").GetBoolean());
    }

    [Fact]
    public void Set_WithSessionId_ShouldSetNotesInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_set_notes.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, notes: "Session speaker notes");
        Assert.Contains("Notes set", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var notesSlide = ppt.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);

        var isEvaluationMode = IsEvaluationMode();
        if (!isEvaluationMode)
            Assert.Contains("Session speaker notes", notesSlide.NotesTextFrame.Text);
        else
            Assert.False(string.IsNullOrEmpty(notesSlide.NotesTextFrame.Text));
    }

    [Fact]
    public void Clear_WithSessionId_ShouldClearInMemory()
    {
        var pptPath = CreatePresentationWithNotes("test_session_clear_notes.pptx", "Notes to clear in session");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("clear", sessionId: sessionId, slideIndices: [0]);
        Assert.Contains("Cleared", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var notesSlide = ppt.Slides[0].NotesSlideManager.NotesSlide;
        Assert.True(string.IsNullOrEmpty(notesSlide?.NotesTextFrame?.Text));
    }

    #endregion
}