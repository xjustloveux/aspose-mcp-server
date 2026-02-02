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
    public void Constructor_WithValidParameters_ShouldStoreFields()
    {
        var handler = new WebSocketConnectionHandler("/path/to/executable", "--arg1");

        Assert.NotNull(handler);
        var executableField = typeof(WebSocketConnectionHandler)
            .GetField("_executablePath", BindingFlags.NonPublic | BindingFlags.Instance);
        var toolArgsField = typeof(WebSocketConnectionHandler)
            .GetField("_toolArgs", BindingFlags.NonPublic | BindingFlags.Instance);
        Assert.Equal("/path/to/executable", executableField?.GetValue(handler));
        Assert.Equal("--arg1", toolArgsField?.GetValue(handler));
    }

    [Fact]
    public void Constructor_WithEmptyExecutablePath_ShouldStoreEmptyValues()
    {
        var handler = new WebSocketConnectionHandler("", "");

        Assert.NotNull(handler);
        var executableField = typeof(WebSocketConnectionHandler)
            .GetField("_executablePath", BindingFlags.NonPublic | BindingFlags.Instance);
        var toolArgsField = typeof(WebSocketConnectionHandler)
            .GetField("_toolArgs", BindingFlags.NonPublic | BindingFlags.Instance);
        Assert.Equal("", executableField?.GetValue(handler));
        Assert.Equal("", toolArgsField?.GetValue(handler));
    }

    [Fact]
    public void Constructor_WithLoggerFactory_ShouldCreateLogger()
    {
        var mockLoggerFactory = new Mock<ILoggerFactory>();
        var mockLogger = new Mock<ILogger<WebSocketConnectionHandler>>();
        mockLoggerFactory
            .Setup(x => x.CreateLogger(It.IsAny<string>()))
            .Returns(mockLogger.Object);

        var handler = new WebSocketConnectionHandler("/path/to/exe", "--args", mockLoggerFactory.Object);

        Assert.NotNull(handler);
        var loggerField = typeof(WebSocketConnectionHandler)
            .GetField("_logger", BindingFlags.NonPublic | BindingFlags.Instance);
        Assert.NotNull(loggerField?.GetValue(handler));
        mockLoggerFactory.Verify(x => x.CreateLogger(It.IsAny<string>()), Times.Once);
    }

    [Fact]
    public void Constructor_WithNullLoggerFactory_ShouldHaveNullLogger()
    {
        var handler = new WebSocketConnectionHandler("/path/to/exe", "--args");

        Assert.NotNull(handler);
        var loggerField = typeof(WebSocketConnectionHandler)
            .GetField("_logger", BindingFlags.NonPublic | BindingFlags.Instance);
        Assert.Null(loggerField?.GetValue(handler));
    }

    #endregion
}
