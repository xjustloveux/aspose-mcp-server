using System.ComponentModel;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Email;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Email;

/// <summary>
///     Tool for converting emails between different formats.
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Email.Conversion")]
[McpServerToolType]
public class EmailConversionTool
{
    /// <summary>
    ///     Handler registry for email conversion operations.
    /// </summary>
    private readonly HandlerRegistry<object> _handlerRegistry;

    /// <summary>
    ///     Server configuration for path-allowlist enforcement.
    /// </summary>
    private readonly ServerConfig? _serverConfig;

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailConversionTool" /> class.
    /// </summary>
    /// <param name="serverConfig">Optional server config for path allowlist enforcement.</param>
    public EmailConversionTool(ServerConfig? serverConfig = null)
    {
        _serverConfig = serverConfig;
        _handlerRegistry =
            HandlerRegistry<object>.CreateFromNamespace("AsposeMcpServer.Handlers.Email.Conversion");
    }

    /// <summary>
    ///     Executes an email conversion operation.
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: convert.
    /// </param>
    /// <param name="path">Source email file path (EML, MSG, etc.).</param>
    /// <param name="outputPath">Output file path with the target format extension.</param>
    /// <returns>An email conversion result containing conversion details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "email_conversion",
        Title = "Email Format Conversion",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Convert emails between different formats. Supports 1 operation: convert.

Usage examples:
- EML to MSG: email_conversion(operation='convert', path='email.eml', outputPath='email.msg')
- MSG to HTML: email_conversion(operation='convert', path='email.msg', outputPath='email.html')
- EML to MHT: email_conversion(operation='convert', path='email.eml', outputPath='email.mht')
- MSG to EML: email_conversion(operation='convert', path='email.msg', outputPath='email.eml')

Supported formats: EML, EMLX, MSG, MHT/MHTML, HTML")]
    public object Execute(
        [Description(@"Operation to perform.
- 'convert': Convert email between formats (required params: path, outputPath)")]
        string operation,
        [Description("Source email file path (EML, EMLX, MSG, MHT, HTML)")]
        string? path = null,
        [Description("Output file path (format determined by extension: .eml, .emlx, .msg, .mht, .mhtml, .html, .htm)")]
        string? outputPath = null)
    {
        var parameters = BuildParameters(path, outputPath);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<object>
        {
            Document = new object(),
            SourcePath = path,
            OutputPath = outputPath,
            ServerConfig = _serverConfig
        };

        object result;
        try
        {
            result = handler.Execute(operationContext, parameters);
        }
        catch (Exception ex) when (ex is not ArgumentException and not KeyNotFoundException
                                       and not FileNotFoundException and not UnauthorizedAccessException)
        {
            // Aspose.Email failures reached the caller as the global sentinel, which said nothing
            // about what went wrong. The family translator turns the ones it recognises into an
            // actionable message and still sanitises everything else.
            throw EmailErrorTranslator.Translate(ex, Path.GetFileName(path ?? string.Empty));
        }

        return ResultHelper.FinalizeResult((dynamic)result, outputPath, (string?)null);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="path">The source email file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <returns>OperationParameters configured for the conversion operation.</returns>
    private static OperationParameters BuildParameters(string? path, string? outputPath)
    {
        var parameters = new OperationParameters();
        parameters.SetIfNotNull("path", path);
        parameters.SetIfNotNull("outputPath", outputPath);
        return parameters;
    }
}
