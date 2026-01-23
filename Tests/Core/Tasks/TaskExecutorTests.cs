using AsposeMcpServer.Core.Tasks;

namespace AsposeMcpServer.Tests.Core.Tasks;

public class TaskExecutorTests
{
    [Theory]
    [InlineData("convert_to_pdf")]
    [InlineData("convert_document")]
    [InlineData("CONVERT_TO_PDF")]
    [InlineData("Convert_Document")]
    public void SupportsAsync_WithSupportedTool_ShouldReturnTrue(string toolName)
    {
        Assert.True(TaskExecutor.SupportsAsync(toolName));
    }

    [Theory]
    [InlineData("word_text")]
    [InlineData("excel_cell")]
    [InlineData("pdf_text")]
    [InlineData("unknown_tool")]
    public void SupportsAsync_WithUnsupportedTool_ShouldReturnFalse(string toolName)
    {
        Assert.False(TaskExecutor.SupportsAsync(toolName));
    }

    [Fact]
    public void SupportedTools_ShouldContainExpectedTools()
    {
        Assert.Contains("convert_to_pdf", TaskExecutor.SupportedTools);
        Assert.Contains("convert_document", TaskExecutor.SupportedTools);
        Assert.Equal(2, TaskExecutor.SupportedTools.Count);
    }

    [Fact]
    public void Constructor_WithNullStore_ShouldThrow()
    {
        var services = new MockServiceProvider();

        Assert.Throws<ArgumentNullException>(() => new TaskExecutor(null!, services));
    }

    [Fact]
    public void Constructor_WithNullServices_ShouldThrow()
    {
        var config = new TaskConfig();
        var store = new TaskStore(config);

        Assert.Throws<ArgumentNullException>(() => new TaskExecutor(store, null!));
    }

    private class MockServiceProvider : IServiceProvider
    {
        public object? GetService(Type serviceType)
        {
            return null;
        }
    }
}
