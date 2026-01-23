using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Base class for Handler tests providing common test infrastructure.
/// </summary>
/// <typeparam name="TContext">The document context type.</typeparam>
public abstract class HandlerTestBase<TContext> : TestBase where TContext : class
{
    /// <summary>
    ///     Creates a temporary file in the test directory with the specified extension and content.
    ///     The file is automatically cleaned up when the test completes.
    /// </summary>
    /// <param name="extension">The file extension (e.g., ".txt", ".bmp").</param>
    /// <param name="content">The text content to write.</param>
    /// <returns>The full path to the created file.</returns>
    protected string CreateTempFile(string extension, string content)
    {
        var tempPath = Path.Combine(TestDir, $"temp_{Guid.NewGuid()}{extension}");
        File.WriteAllText(tempPath, content);
        return tempPath;
    }

    /// <summary>
    ///     Creates a temporary file in the test directory with the specified extension and binary content.
    ///     The file is automatically cleaned up when the test completes.
    /// </summary>
    /// <param name="extension">The file extension (e.g., ".bmp", ".png").</param>
    /// <param name="content">The binary content to write.</param>
    /// <returns>The full path to the created file.</returns>
    protected string CreateTempFile(string extension, byte[] content)
    {
        var tempPath = Path.Combine(TestDir, $"temp_{Guid.NewGuid()}{extension}");
        File.WriteAllBytes(tempPath, content);
        return tempPath;
    }

    /// <summary>
    ///     Creates a simple BMP image file for testing.
    ///     The file is automatically cleaned up when the test completes.
    /// </summary>
    /// <returns>The full path to the created image file.</returns>
    protected string CreateTempImageFile()
    {
        var width = 10;
        var height = 10;
        var bmp = new byte[width * height * 3 + 54];
        bmp[0] = 0x42;
        bmp[1] = 0x4D;
        var fileSize = bmp.Length;
        bmp[2] = (byte)(fileSize & 0xFF);
        bmp[3] = (byte)((fileSize >> 8) & 0xFF);
        bmp[4] = (byte)((fileSize >> 16) & 0xFF);
        bmp[5] = (byte)((fileSize >> 24) & 0xFF);
        bmp[10] = 54;
        bmp[14] = 40;
        bmp[18] = (byte)(width & 0xFF);
        bmp[19] = (byte)((width >> 8) & 0xFF);
        bmp[22] = (byte)(height & 0xFF);
        bmp[23] = (byte)((height >> 8) & 0xFF);
        bmp[26] = 1;
        bmp[28] = 24;
        for (var i = 54; i < bmp.Length; i += 3)
        {
            bmp[i] = 255;
            bmp[i + 1] = 0;
            bmp[i + 2] = 0;
        }

        return CreateTempFile(".bmp", bmp);
    }

    /// <summary>
    ///     Creates an operation context for testing.
    /// </summary>
    /// <param name="document">The document instance.</param>
    /// <param name="outputPath">Optional output path.</param>
    /// <returns>The operation context.</returns>
    protected static OperationContext<TContext> CreateContext(TContext document, string? outputPath = null)
    {
        return new OperationContext<TContext>
        {
            Document = document,
            OutputPath = outputPath
        };
    }

    /// <summary>
    ///     Creates operation parameters from a dictionary.
    /// </summary>
    /// <param name="values">The parameter values.</param>
    /// <returns>The operation parameters.</returns>
    protected static OperationParameters CreateParameters(Dictionary<string, object?> values)
    {
        var parameters = new OperationParameters();
        foreach (var (key, value) in values)
            parameters.Set(key, value);
        return parameters;
    }

    /// <summary>
    ///     Creates empty operation parameters.
    /// </summary>
    /// <returns>Empty operation parameters.</returns>
    protected static OperationParameters CreateEmptyParameters()
    {
        return new OperationParameters();
    }

    /// <summary>
    ///     Asserts that the handler execution marks the context as modified.
    /// </summary>
    /// <param name="context">The operation context.</param>
    protected static void AssertModified(OperationContext<TContext> context)
    {
        Assert.True(context.IsModified, "Context should be marked as modified");
    }

    /// <summary>
    ///     Asserts that the handler execution does not mark the context as modified.
    /// </summary>
    /// <param name="context">The operation context.</param>
    protected static void AssertNotModified(OperationContext<TContext> context)
    {
        Assert.False(context.IsModified, "Context should not be marked as modified");
    }
}
