using AsposeMcpServer.Core.Extension;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Tests for CommandResponse class.
/// </summary>
public class CommandResponseTests
{
    [Fact]
    public void Success_WithCommandId_SetsProperties()
    {
        var response = CommandResponse.Success("cmd123");

        Assert.True(response.IsSuccess);
        Assert.Equal("cmd123", response.CommandId);
        Assert.Null(response.Error);
    }

    [Fact]
    public void Success_WithResult_SetsProperties()
    {
        var result = new Dictionary<string, object> { { "key", "value" } };
        var response = CommandResponse.Success("cmd123", result);

        Assert.True(response.IsSuccess);
        Assert.NotNull(response.Result);
        Assert.Equal("value", response.Result["key"]);
    }

    [Fact]
    public void Failure_WithError_SetsProperties()
    {
        var response = CommandResponse.Failure("Something went wrong");

        Assert.False(response.IsSuccess);
        Assert.Equal("Something went wrong", response.Error);
    }

    [Fact]
    public void Failure_WithCommandId_SetsProperties()
    {
        var response = CommandResponse.Failure("Error message", "cmd456");

        Assert.False(response.IsSuccess);
        Assert.Equal("cmd456", response.CommandId);
        Assert.Equal("Error message", response.Error);
    }
}
