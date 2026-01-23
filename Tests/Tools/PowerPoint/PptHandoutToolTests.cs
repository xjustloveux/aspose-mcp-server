using Aspose.Slides;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptHandoutTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptHandoutToolTests : PptTestBase
{
    private readonly PptHandoutTool _tool;

    public PptHandoutToolTests()
    {
        _tool = new PptHandoutTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void SetHeaderFooter_WithoutHandoutMaster_ShouldThrowWithHelpfulMessage()
    {
        var pptPath = CreatePresentation("test_handout_no_master.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath, headerText: "Handout Header"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("SET_HEADER_FOOTER")]
    [InlineData("Set_Header_Footer")]
    [InlineData("set_header_footer")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_{operation.Replace("_", "")}.pptx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute(operation, pptPath, headerText: "Test"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
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
    public void SetHeaderFooter_WithSessionId_WithoutHandoutMaster_ShouldThrow()
    {
        var pptPath = CreatePresentation("test_session_set_header_footer.pptx");
        var sessionId = OpenSession(pptPath);
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", sessionId: sessionId, headerText: "Session Header"));
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
        var pptPath1 = CreatePresentation("test_path_handout.pptx");
        var pptPath2 = CreatePresentation("test_session_handout.pptx");
        var sessionId = OpenSession(pptPath2);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotNull(ppt);
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("set_header_footer", pptPath1, sessionId, headerText: "Test"));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
