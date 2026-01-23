using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Properties;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptPropertiesTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptPropertiesToolTests : PptTestBase
{
    private readonly PptPropertiesTool _tool;

    public PptPropertiesToolTests()
    {
        _tool = new PptPropertiesTool(SessionManager);
    }

    private string CreatePresentationWithProperties(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.DocumentProperties.Title = "Original Title";
        presentation.DocumentProperties.Author = "Original Author";
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Get_ShouldReturnPropertiesAsJson()
    {
        var pptPath = CreatePresentation("test_get.pptx");
        var result = _tool.Execute("get", pptPath);
        var data = GetResultData<GetPropertiesPptResult>(result);
        Assert.NotNull(data);
        Assert.NotNull(data.Title);
        Assert.NotNull(data.Author);
    }

    [Fact]
    public void Set_ShouldSetBasicProperties()
    {
        var pptPath = CreatePresentation("test_set_basic.pptx");
        var outputPath = CreateTestFilePath("test_set_basic_output.pptx");
        var result = _tool.Execute("set", pptPath, title: "Test Presentation", author: "Test Author",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Document properties updated", data.Message);
        Assert.Contains("Title", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_get_{operation}.pptx");
        var result = _tool.Execute(operation, pptPath);
        var data = GetResultData<GetPropertiesPptResult>(result);
        Assert.NotNull(data);
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
    public void Get_WithSessionId_ShouldReturnPropertiesFromMemory()
    {
        var pptPath = CreatePresentationWithProperties("test_session_get.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetPropertiesPptResult>(result);
        Assert.NotNull(data);
        Assert.NotNull(data.Title);
        Assert.NotNull(data.Author);
        var output = GetResultOutput<GetPropertiesPptResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Set_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreatePresentation("test_session_set.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set", sessionId: sessionId, title: "Session Title", author: "Session Author");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Document properties updated", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithProperties("test_path_props.pptx");
        var pptPath2 = CreatePresentation("test_session_props.pptx");
        using (var ppt = new Presentation(pptPath2))
        {
            ppt.DocumentProperties.Title = "SessionTitle";
            ppt.Save(pptPath2, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId);
        var data = GetResultData<GetPropertiesPptResult>(result);
        Assert.NotNull(data);
        var output = GetResultOutput<GetPropertiesPptResult>(result);
        Assert.True(output.IsSession);
    }

    #endregion
}
