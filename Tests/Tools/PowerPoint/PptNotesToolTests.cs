using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Notes;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptNotesTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
[SupportedOSPlatform("windows")]
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

    [SkippableFact]
    public void Set_ShouldSetNotes()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentation("test_set.pptx");
        var outputPath = CreateTestFilePath("test_set_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, notes: "Speaker notes for this slide",
            outputPath: outputPath);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        using var presentation = new Presentation(outputPath);
        var notesSlide = presentation.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);
        if (!IsEvaluationMode())
            Assert.Contains("Speaker notes for this slide", notesSlide.NotesTextFrame.Text);
    }

    [SkippableFact]
    public void Get_WithSlideIndex_ShouldReturnSingleSlideNotes()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithNotes("test_get.pptx", "Test notes content");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        var data = GetResultData<GetNotesResult>(result);
        Assert.Equal(0, data.SlideIndex);
        Assert.True(data.HasNotes);
    }

    [SkippableFact]
    public void Clear_WithSlideIndices_ShouldClearSpecificSlides()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithNotes("test_clear.pptx", "Notes to clear");
        var outputPath = CreateTestFilePath("test_clear_output.pptx");
        var result = _tool.Execute("clear", pptPath, slideIndices: [0], outputPath: outputPath);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void SetHeaderFooter_ShouldSetNotesHeaderFooter()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentation("test_hf.pptx");
        var outputPath = CreateTestFilePath("test_hf_output.pptx");
        var result = _tool.Execute("set_header_footer", pptPath, headerText: "Notes Header", footerText: "Notes Footer",
            outputPath: outputPath);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [SkippableTheory]
    [InlineData("SET")]
    [InlineData("Set")]
    [InlineData("set")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentation($"test_case_set_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_set_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, notes: "Test", outputPath: outputPath);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [SkippableFact]
    public void Get_WithSessionId_ShouldReturnNotes()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithNotes("test_session_get.pptx", "Session test notes");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        var data = GetResultData<GetNotesResult>(result);
        Assert.Equal(0, data.SlideIndex);
        Assert.True(data.HasNotes);
        var output = GetResultOutput<GetNotesResult>(result);
        Assert.True(output.IsSession);
    }

    [SkippableFact]
    public void Set_WithSessionId_ShouldSetNotesInMemory()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentation("test_session_set.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, notes: "Session speaker notes");
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var notesSlide = ppt.Slides[0].NotesSlideManager.NotesSlide;
        Assert.NotNull(notesSlide);
        if (!IsEvaluationMode())
            Assert.Contains("Session speaker notes", notesSlide.NotesTextFrame.Text);
    }

    [SkippableFact]
    public void Clear_WithSessionId_ShouldClearInMemory()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithNotes("test_session_clear.pptx", "Notes to clear in session");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("clear", sessionId: sessionId, slideIndices: [0]);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [SkippableFact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        SkipIfNotWindows();
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session", slideIndex: 0));
    }

    [SkippableFact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        SkipIfNotWindows();
        var pptPath1 = CreatePresentationWithNotes("test_path_notes.pptx", "PathNotes");
        var pptPath2 = CreatePresentationWithNotes("test_session_notes.pptx", "SessionNotes");
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId, slideIndex: 0);
        var data = GetResultData<GetNotesResult>(result);
        Assert.Equal(0, data.SlideIndex);
        var output = GetResultOutput<GetNotesResult>(result);
        Assert.True(output.IsSession);
    }

    #endregion
}
