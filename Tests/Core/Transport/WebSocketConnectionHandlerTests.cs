using System.Reflection;
using AsposeMcpServer.Core.Transport;
using Microsoft.Extensions.Logging;
using Moq;

namespace AsposeMcpServer.Tests.Core.Transport;

public class WebSocketConnectionHandlerTests
{
    #region MaxMessageSize Tests

    [Fact]
    public void MaxMessageSize_ShouldBe10MB()
    {
        var expectedSize = 10 * 1024 * 1024;
        var field = typeof(WebSocketConnectionHandler)
            .GetField("MaxMessageSize", BindingFlags.NonPublic | BindingFlags.Static);

        var actualSize = field?.GetValue(null);

        Assert.Equal(expectedSize, actualSize);
    }

    #endregion

    #region Constructor Tests

    [Fact]
    public void Constructor_WithValidParameters_ShouldNotThrow()
    {
        var handler = new WebSocketConnectionHandler("/path/to/executable", "--arg1");

        Assert.NotNull(handler);
    }

    [Fact]
    public void Constructor_WithEmptyExecutablePath_ShouldNotThrow()
    {
        var handler = new WebSocketConnectionHandler("", "");

        Assert.NotNull(handler);
    }

    [Fact]
    public void Constructor_WithLoggerFactory_ShouldNotThrow()
    {
        var mockLoggerFactory = new Mock<ILoggerFactory>();
        var mockLogger = new Mock<ILogger<WebSocketConnectionHandler>>();
        mockLoggerFactory
            .Setup(x => x.CreateLogger(It.IsAny<string>()))
            .Returns(mockLogger.Object);

        var handler = new WebSocketConnectionHandler("/path/to/exe", "--args", mockLoggerFactory.Object);

        Assert.NotNull(handler);
    }

    [Fact]
    public void Constructor_WithNullLoggerFactory_ShouldNotThrow()
    {
        var handler = new WebSocketConnectionHandler("/path/to/exe", "--args");

        Assert.NotNull(handler);
    }

    #endregion
}
