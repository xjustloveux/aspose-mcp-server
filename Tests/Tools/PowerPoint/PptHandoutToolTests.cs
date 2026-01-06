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

    #region General

    [Fact]
    public void SetHeaderFooter_WithoutHandoutMaster_ShouldThrowWithHelpfulMessage()
    {
        var pptPath = CreateTestPresentation("test_handout_no_master.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath, headerText: "Handout Header"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("PowerPoint", ex.Message);
        Assert.Contains("View", ex.Message);
        Assert.Contains("Handout Master", ex.Message);
    }

    [Fact]
    public void SetHeaderFooter_WithAllParameters_ShouldThrowWithoutMaster()
    {
        var pptPath = CreateTestPresentation("test_handout_all_params.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() => _tool.Execute("set_header_footer", pptPath,
            headerText: "Header", footerText: "Footer", dateText: "2024-12-28", showPageNumber: true));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("SET_HEADER_FOOTER")]
    [InlineData("Set_Header_Footer")]
    [InlineData("set_header_footer")]
    public void Operation_ShouldBeCaseInsensitive_SetHeaderFooter(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_{operation.Replace("_", "")}.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute(operation, pptPath, headerText: "Test"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
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
    public void SetHeaderFooter_WithOnlyFooter_ShouldThrowWithoutMaster()
    {
        var pptPath = CreateTestPresentation("test_footer_only.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath, footerText: "Footer Only"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetHeaderFooter_WithOnlyDate_ShouldThrowWithoutMaster()
    {
        var pptPath = CreateTestPresentation("test_date_only.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath, dateText: "2024-12-28"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetHeaderFooter_WithPageNumberOnly_ShouldThrowWithoutMaster()
    {
        var pptPath = CreateTestPresentation("test_pagenumber_only.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath, showPageNumber: true));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session

    [Fact]
    public void SetHeaderFooter_WithSessionId_WithoutHandoutMaster_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_session_set_header_footer.pptx");
        var sessionId = OpenSession(pptPath);
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", sessionId: sessionId, headerText: "Session Header"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithSessionId_ShouldAccessDocument()
    {
        var pptPath = CreateTestPresentation("test_session_access.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotNull(ppt);
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", sessionId: sessionId, footerText: "Test Footer"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("set_header_footer", sessionId: "invalid_session", headerText: "Test"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreateTestPresentation("test_path_handout.pptx");
        var pptPath2 = CreateTestPresentation("test_session_handout.pptx");
        var sessionId = OpenSession(pptPath2);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotNull(ppt);
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath1, sessionId, headerText: "Test"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}