using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptHandoutToolTests : TestBase
{
    private readonly PptHandoutTool _tool;

    public PptHandoutToolTests()
    {
        _tool = new PptHandoutTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void SetHeaderFooter_WithoutHandoutMaster_ShouldThrowWithHelpfulMessage()
    {
        // Arrange - New presentations don't have handout master by default
        var pptPath = CreateTestPresentation("test_handout_no_master.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath, headerText: "Handout Header"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("PowerPoint", ex.Message);
    }

    [Fact]
    public void SetHeaderFooter_ErrorMessage_ShouldContainInstructions()
    {
        // Arrange - Verify the error message contains helpful instructions
        var pptPath = CreateTestPresentation("test_handout_instructions.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath, headerText: "Header"));

        // Assert - Error message should contain instructions for creating handout master
        Assert.Contains("View", ex.Message);
        Assert.Contains("Handout Master", ex.Message);
    }

    [Fact]
    public void SetHeaderFooter_WithAllParameters_ShouldThrowWithoutMaster()
    {
        // Arrange - Test with all parameters to ensure they are parsed correctly before the error
        var pptPath = CreateTestPresentation("test_handout_all_params.pptx");

        // Act & Assert - Should throw because no handout master, but parameters should be valid
        var ex = Assert.Throws<InvalidOperationException>(() => _tool.Execute("set_header_footer", pptPath,
            headerText: "Header", footerText: "Footer", dateText: "2024-12-28", showPageNumber: true));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_UnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
    }

    [Fact]
    public void SetHeaderFooter_NoHandoutMaster_ShouldThrowInvalidOperationException()
    {
        var pptPath = CreateTestPresentation("test_no_text_params.pptx");

        // Act & Assert - Should throw because no handout master exists in the presentation
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath, headerText: "Test"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void SetHeaderFooter_WithSessionId_WithoutHandoutMaster_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_session_set_header_footer.pptx");
        var sessionId = OpenSession(pptPath);

        // Act & Assert - Should throw because no handout master
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", sessionId: sessionId, headerText: "Session Header"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithSessionId_ShouldAccessDocument()
    {
        var pptPath = CreateTestPresentation("test_session_access.pptx");
        var sessionId = OpenSession(pptPath);

        // Verify we can access the document in memory
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotNull(ppt);

        // Act & Assert - Using set_header_footer without handout master should throw
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", sessionId: sessionId, footerText: "Test Footer"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}