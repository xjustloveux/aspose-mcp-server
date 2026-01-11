using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptNotesTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptNotesToolTests : PptTestBase
{
    private readonly PptNotesTool _tool;

    public PptNotesToolTests()
    {
        _tool = new PptNotesTool(SessionManager);
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

    #region File I/O Smoke Tests

    [Fact]
    public void Set_ShouldSetNotes()
    {
        var pptPath = CreatePresentation("test_set.pptx");
        var outputPath = CreateTestFilePath("test_set_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, notes: "Speaker notes for this slide",
            outputPath: outputPath);
        Assert.StartsWith("Notes set for slide", result);
        using var presentation = new Presentation(outputPath);
        var notesSlide = presentation.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);
    }

    [Fact]
    public void Get_WithSlideIndex_ShouldReturnSingleSlideNotes()
    {
        var pptPath = CreatePresentationWithNotes("test_get.pptx", "Test notes content");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("slideIndex").GetInt32());
        Assert.True(json.RootElement.GetProperty("hasNotes").GetBoolean());
    }

    [Fact]
    public void Clear_WithSlideIndices_ShouldClearSpecificSlides()
    {
        var pptPath = CreatePresentationWithNotes("test_clear.pptx", "Notes to clear");
        var outputPath = CreateTestFilePath("test_clear_output.pptx");
        var result = _tool.Execute("clear", pptPath, slideIndices: [0], outputPath: outputPath);
        Assert.StartsWith("Cleared speaker notes for", result);
    }

    [Fact]
    public void SetHeaderFooter_ShouldSetNotesHeaderFooter()
    {
        var pptPath = CreatePresentation("test_hf.pptx");
        var outputPath = CreateTestFilePath("test_hf_output.pptx");
        var result = _tool.Execute("set_header_footer", pptPath, headerText: "Notes Header", footerText: "Notes Footer",
            outputPath: outputPath);
        Assert.StartsWith("Notes master header/footer updated", result);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("SET")]
    [InlineData("Set")]
    [InlineData("set")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_set_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_set_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, notes: "Test", outputPath: outputPath);
        Assert.StartsWith("Notes set for slide", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

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
        var pptPath = CreatePresentation("test_session_set.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, notes: "Session speaker notes");
        Assert.StartsWith("Notes set for slide", result);
        var notesSlide = ppt.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);
    }

    [Fact]
    public void Clear_WithSessionId_ShouldClearInMemory()
    {
        var pptPath = CreatePresentationWithNotes("test_session_clear.pptx", "Notes to clear in session");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("clear", sessionId: sessionId, slideIndices: [0]);
        Assert.StartsWith("Cleared speaker notes for", result);
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
        Assert.Contains("slideIndex", result);
    }

    #endregion
}
