using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Session;
using AsposeMcpServer.Tools.Session;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Integration.ErrorHandling;

/// <summary>
///     Integration tests for parameter validation error handling.
/// </summary>
[Trait("Category", "Integration")]
public class ValidationErrorTests : IntegrationTestBase
{
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;
    private readonly WordTextTool _textTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ValidationErrorTests" /> class.
    /// </summary>
    public ValidationErrorTests()
    {
        var config = new SessionConfig { Enabled = true };
        _sessionManager = new DocumentSessionManager(config);
        var tempFileManager = new TempFileManager(config);
        _sessionTool = new DocumentSessionTool(_sessionManager, tempFileManager, new StdioSessionIdentityAccessor());
        _textTool = new WordTextTool(_sessionManager);
    }

    /// <summary>
    ///     Disposes of test resources.
    /// </summary>
    public override void Dispose()
    {
        _sessionManager.Dispose();
        base.Dispose();
    }

    #region Empty Parameter Tests

    /// <summary>
    ///     Verifies that empty find parameter throws ArgumentException.
    /// </summary>
    [Fact]
    public void Validation_EmptyFindText_ThrowsArgumentException()
    {
        var path = CreateWordDocument();
        var openResult = _sessionTool.Execute("open", path);
        var openData = GetResultData<OpenSessionResult>(openResult);

        Assert.Throws<ArgumentException>(() =>
            _textTool.Execute("replace", sessionId: openData.SessionId, find: "", replace: "New"));
    }

    #endregion

    #region Missing Required Parameter Tests

    /// <summary>
    ///     Verifies that missing find parameter throws ArgumentException.
    /// </summary>
    [Fact]
    public void Validation_MissingFindText_ThrowsArgumentException()
    {
        var path = CreateWordDocument();
        var openResult = _sessionTool.Execute("open", path);
        var openData = GetResultData<OpenSessionResult>(openResult);

        Assert.Throws<ArgumentException>(() =>
            _textTool.Execute("replace", sessionId: openData.SessionId, replace: "New"));
    }

    /// <summary>
    ///     Verifies that missing replace parameter throws ArgumentException.
    /// </summary>
    [Fact]
    public void Validation_MissingReplaceText_ThrowsArgumentException()
    {
        var path = CreateWordDocument();
        var openResult = _sessionTool.Execute("open", path);
        var openData = GetResultData<OpenSessionResult>(openResult);

        Assert.Throws<ArgumentException>(() =>
            _textTool.Execute("replace", sessionId: openData.SessionId, find: "Old"));
    }

    #endregion

    #region Invalid Operation Tests

    /// <summary>
    ///     Verifies that an invalid operation throws ArgumentException.
    /// </summary>
    [Fact]
    public void Validation_InvalidOperation_ThrowsArgumentException()
    {
        var path = CreateWordDocument();

        Assert.Throws<ArgumentException>(() =>
            _sessionTool.Execute("invalid_operation", path));
    }

    /// <summary>
    ///     Verifies that an unknown text operation throws ArgumentException.
    /// </summary>
    [Fact]
    public void Validation_UnknownTextOperation_ThrowsArgumentException()
    {
        var path = CreateWordDocument();
        var openResult = _sessionTool.Execute("open", path);
        var openData = GetResultData<OpenSessionResult>(openResult);

        Assert.Throws<ArgumentException>(() =>
            _textTool.Execute("unknown_operation", sessionId: openData.SessionId));
    }

    #endregion
}
