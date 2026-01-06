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

    private string CreateEmptyPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.RemoveAt(0);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General

    [Fact]
    public void Set_ShouldSetNotes()
    {
        var pptPath = CreateTestPresentation("test_set.pptx");
        var outputPath = CreateTestFilePath("test_set_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, notes: "Speaker notes for this slide",
            outputPath: outputPath);
        Assert.StartsWith("Notes set for slide", result);
        using var presentation = new Presentation(outputPath);
        var notesSlide = presentation.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);
        Assert.False(string.IsNullOrEmpty(notesSlide.NotesTextFrame.Text));
        if (!IsEvaluationMode())
            Assert.Contains("Speaker notes", notesSlide.NotesTextFrame.Text);
        else
            // Fallback: verify basic structure in evaluation mode
            Assert.NotNull(notesSlide.NotesTextFrame);
    }

    [Fact]
    public void Set_ShouldReplaceExistingNotes()
    {
        var pptPath = CreatePresentationWithNotes("test_set_replace.pptx", "Original notes");
        var outputPath = CreateTestFilePath("test_set_replace_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, notes: "Updated notes content",
            outputPath: outputPath);
        Assert.StartsWith("Notes set for slide", result);
        using var presentation = new Presentation(outputPath);
        var notesSlide = presentation.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);
        if (IsEvaluationMode())
            Assert.DoesNotContain("Original notes", notesSlide.NotesTextFrame.Text);
        else
            Assert.Equal("Updated notes content", notesSlide.NotesTextFrame.Text);
    }

    [Fact]
    public void Get_WithSlideIndex_ShouldReturnSingleSlideNotes()
    {
        var pptPath = CreatePresentationWithNotes("test_get.pptx", "Test notes content");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("slideIndex").GetInt32());
        Assert.True(json.RootElement.GetProperty("hasNotes").GetBoolean());
        Assert.False(string.IsNullOrEmpty(json.RootElement.GetProperty("notes").GetString()));
        if (!IsEvaluationMode())
            Assert.Contains("Test notes", json.RootElement.GetProperty("notes").GetString());
        else
            // Fallback: verify basic structure in evaluation mode
            Assert.True(json.RootElement.TryGetProperty("notes", out _));
    }

    [Fact]
    public void Get_WithoutSlideIndex_ShouldReturnAllNotes()
    {
        var pptPath = CreateTestPresentation("test_get_all.pptx", 3);
        var result = _tool.Execute("get", pptPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(3, json.RootElement.GetProperty("slides").GetArrayLength());
    }

    [Fact]
    public void Clear_WithSlideIndices_ShouldClearSpecificSlides()
    {
        var pptPath = CreatePresentationWithNotes("test_clear.pptx", "Notes to clear");
        var outputPath = CreateTestFilePath("test_clear_output.pptx");
        var result = _tool.Execute("clear", pptPath, slideIndices: [0], outputPath: outputPath);
        Assert.StartsWith("Cleared speaker notes for", result);
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
        Assert.StartsWith("Cleared speaker notes for", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetHeaderFooter_ShouldSetNotesHeaderFooter()
    {
        var pptPath = CreateTestPresentation("test_hf.pptx");
        var outputPath = CreateTestFilePath("test_hf_output.pptx");
        var result = _tool.Execute("set_header_footer", pptPath, headerText: "Notes Header", footerText: "Notes Footer",
            outputPath: outputPath);
        Assert.StartsWith("Notes master header/footer updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetHeaderFooter_WithAllParameters_ShouldSetAll()
    {
        var pptPath = CreateTestPresentation("test_hf_all.pptx");
        var outputPath = CreateTestFilePath("test_hf_all_output.pptx");
        var result = _tool.Execute("set_header_footer", pptPath, headerText: "Header", footerText: "Footer",
            dateText: "2024-12-28", showPageNumber: true, outputPath: outputPath);
        Assert.StartsWith("Notes master header/footer updated", result);
    }

    [Fact]
    public void SetHeaderFooter_HidePageNumber_ShouldHide()
    {
        var pptPath = CreateTestPresentation("test_hf_hide.pptx");
        var outputPath = CreateTestFilePath("test_hf_hide_output.pptx");
        var result = _tool.Execute("set_header_footer", pptPath, showPageNumber: false, outputPath: outputPath);
        Assert.StartsWith("Notes master header/footer updated", result);
    }

    [Theory]
    [InlineData("SET")]
    [InlineData("Set")]
    [InlineData("set")]
    public void Operation_ShouldBeCaseInsensitive_Set(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_set_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_set_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, notes: "Test", outputPath: outputPath);
        Assert.StartsWith("Notes set for slide", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_get_{operation}.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0);
        Assert.Contains("slideIndex", result);
    }

    [Theory]
    [InlineData("CLEAR")]
    [InlineData("Clear")]
    [InlineData("clear")]
    public void Operation_ShouldBeCaseInsensitive_Clear(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_clear_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_clear_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, outputPath: outputPath);
        Assert.StartsWith("Cleared speaker notes for", result);
    }

    [Theory]
    [InlineData("SET_HEADER_FOOTER")]
    [InlineData("Set_Header_Footer")]
    [InlineData("set_header_footer")]
    public void Operation_ShouldBeCaseInsensitive_SetHeaderFooter(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_hf_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_hf_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, headerText: "Header", outputPath: outputPath);
        Assert.StartsWith("Notes master header/footer updated", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Set_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_set_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set", pptPath, notes: "Test"));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Fact]
    public void Set_WithoutNotes_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_set_no_notes.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set", pptPath, slideIndex: 0));
        Assert.Contains("notes is required", ex.Message);
    }

    [Fact]
    public void Set_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_set_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set", pptPath, slideIndex: 99, notes: "Test notes"));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Get_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_get_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get", pptPath, slideIndex: 99));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Clear_WithInvalidSlideIndices_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_clear_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("clear", pptPath, slideIndices: [0, 99]));
        Assert.Contains("Invalid slide indices", ex.Message);
    }

    [Fact]
    public void SetHeaderFooter_OnEmptyPresentation_ShouldThrowInvalidOperationException()
    {
        var pptPath = CreateEmptyPresentation("test_hf_empty.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath, headerText: "Header"));
        Assert.Contains("no slides", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void Get_WithSessionId_ShouldReturnNotes()
    {
        var pptPath = CreatePresentationWithNotes("test_session_get.pptx", "Session test notes");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("slideIndex").GetInt32());
        Assert.True(json.RootElement.GetProperty("hasNotes").GetBoolean());
    }

    [Fact]
    public void Set_WithSessionId_ShouldSetNotesInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_set.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, notes: "Session speaker notes");
        Assert.StartsWith("Notes set for slide", result);
        var notesSlide = ppt.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);
        Assert.False(string.IsNullOrEmpty(notesSlide.NotesTextFrame.Text));
        if (!IsEvaluationMode())
            Assert.Contains("Session speaker notes", notesSlide.NotesTextFrame.Text);
        else
            // Fallback: verify basic structure in evaluation mode
            Assert.NotNull(notesSlide.NotesTextFrame);
    }

    [Fact]
    public void Clear_WithSessionId_ShouldClearInMemory()
    {
        var pptPath = CreatePresentationWithNotes("test_session_clear.pptx", "Notes to clear in session");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("clear", sessionId: sessionId, slideIndices: [0]);
        Assert.StartsWith("Cleared speaker notes for", result);
        var notesSlide = ppt.Slides[0].NotesSlideManager.NotesSlide;
        Assert.True(string.IsNullOrEmpty(notesSlide?.NotesTextFrame?.Text));
    }

    [Fact]
    public void SetHeaderFooter_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_hf.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_header_footer", sessionId: sessionId, headerText: "Session Header");
        Assert.StartsWith("Notes master header/footer updated", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session", slideIndex: 0));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithNotes("test_path_notes.pptx", "PathNotes");
        var pptPath2 = CreatePresentationWithNotes("test_session_notes.pptx", "SessionNotes");
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId, slideIndex: 0);
        if (!IsEvaluationMode())
        {
            Assert.Contains("SessionNotes", result);
            Assert.DoesNotContain("PathNotes", result);
        }
        else
        {
            // Fallback: verify basic structure in evaluation mode
            Assert.NotNull(result);
            Assert.Contains("slideIndex", result);
        }
    }

    #endregion
}