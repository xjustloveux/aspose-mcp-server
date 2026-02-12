using System.ComponentModel;
using System.Text.Json;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Extension;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Extension;

/// <summary>
///     Tool for managing extension bindings and status.
/// </summary>
[McpServerToolType]
public class ExtensionTool
{
    /// <summary>
    ///     The extension manager for managing extension lifecycle.
    /// </summary>
    private readonly ExtensionManager _extensionManager;

    /// <summary>
    ///     The accessor for retrieving current session identity.
    /// </summary>
    private readonly ISessionIdentityAccessor _identityAccessor;

    /// <summary>
    ///     The session bridge for managing session-extension bindings.
    /// </summary>
    private readonly ExtensionSessionBridge _sessionBridge;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExtensionTool" /> class.
    /// </summary>
    /// <param name="extensionManager">The extension manager instance.</param>
    /// <param name="sessionBridge">The session bridge instance.</param>
    /// <param name="identityAccessor">The session identity accessor instance.</param>
    public ExtensionTool(
        ExtensionManager extensionManager,
        ExtensionSessionBridge sessionBridge,
        ISessionIdentityAccessor identityAccessor)
    {
        _extensionManager = extensionManager;
        _sessionBridge = sessionBridge;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Manages extension bindings for real-time document preview.
    /// </summary>
    /// <param name="operation">Operation to perform.</param>
    /// <param name="sessionId">Session ID (for bind, unbind, set_format, bindings, command operations).</param>
    /// <param name="extensionId">Extension ID (for bind, unbind, set_format, status, command operations).</param>
    /// <param name="format">Output format (for bind, set_format operations).</param>
    /// <param name="jpegQuality">JPEG quality 1-100 for JPEG image output (default: 90).</param>
    /// <param name="csvSeparator">CSV field separator character (default: comma).</param>
    /// <param name="pdfCompliance">PDF/A compliance level for PDF output (e.g., PDFA1A, PDFA1B).</param>
    /// <param name="dpi">Image DPI resolution 72-600 for image output (default: 150).</param>
    /// <param name="commandType">Command type (for command operation).</param>
    /// <param name="payload">JSON payload (for command operation).</param>
    /// <returns>Operation result object.</returns>
    [McpServerTool(
        Name = "extension",
        Title = "Manage Extensions",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [OutputSchema(typeof(ExtensionResults))]
    [Description(@"Manage extension bindings for real-time document preview and processing.

Extensions are external processes that can receive document snapshots for preview, cloud sync, compliance checks, etc.

Operations:
- 'list': List all registered extensions and their capabilities.
- 'bind': Bind a session to an extension (starts receiving snapshots).
- 'unbind': Unbind a session from an extension (stops receiving snapshots).
- 'set_format': Change the output format for an existing binding.
- 'status': Get detailed status of a specific extension.
- 'bindings': List all bindings for a session.
- 'command': Send a command to an extension.

Usage examples:
- List extensions: extension(operation='list')
- Bind session: extension(operation='bind', sessionId='sess_abc123', extensionId='pdf-viewer', format='pdf')
- Change format: extension(operation='set_format', sessionId='sess_abc123', extensionId='pdf-viewer', format='html')
- Unbind: extension(operation='unbind', sessionId='sess_abc123', extensionId='pdf-viewer')
- Unbind all: extension(operation='unbind', sessionId='sess_abc123')
- Get status: extension(operation='status', extensionId='pdf-viewer')
- List bindings: extension(operation='bindings', sessionId='sess_abc123')
- Send command: extension(operation='command', sessionId='sess_abc123', extensionId='pdf-viewer', commandType='highlight', payload='{""page"":3}')

After binding, the extension will automatically receive document snapshots when the session is modified.")]
    public async Task<object> ExecuteAsync(
        [Description(@"Operation to perform:
- 'list': List all registered extensions
- 'bind': Bind session to extension (required: sessionId, extensionId, format)
- 'unbind': Unbind session from extension (required: sessionId; optional: extensionId)
- 'set_format': Change output format (required: sessionId, extensionId, format)
- 'status': Get extension status (required: extensionId)
- 'bindings': List session bindings (required: sessionId)
- 'command': Send command to extension (required: sessionId, extensionId, commandType; optional: payload)")]
        string operation,
        [Description("Session ID (for bind, unbind, set_format, bindings, command operations)")]
        string? sessionId = null,
        [Description("Extension ID (for bind, unbind, set_format, status, command operations)")]
        string? extensionId = null,
        [Description("Output format: 'pdf', 'html', 'png', etc. (for bind, set_format operations)")]
        string? format = null,
        [Description("JPEG quality 1-100 (default: 90, for bind, set_format with JPEG output)")]
        int jpegQuality = 90,
        [Description("CSV field separator (default: comma, for bind, set_format with CSV output)")]
        string csvSeparator = ",",
        [Description("PDF/A compliance: PDFA1A, PDFA1B, PDFA2A, PDFA2U, PDFA4 (for bind, set_format with PDF output)")]
        string? pdfCompliance = null,
        [Description("Image DPI resolution 72-600 (default: 150, for bind, set_format with image output)")]
        int dpi = 150,
        [Description("Command type to send to the extension (for command operation)")]
        string? commandType = null,
        [Description("JSON payload for the command (for command operation), e.g., '{\"page\":3}'")]
        string? payload = null)
    {
        var options = new ConversionOptions
        {
            JpegQuality = Math.Clamp(jpegQuality, 1, 100),
            CsvSeparator = csvSeparator,
            PdfCompliance = pdfCompliance,
            Dpi = Math.Clamp(dpi, 72, 600)
        };

        object result = operation.ToLowerInvariant() switch
        {
            "list" => ListExtensions(),
            "bind" => await BindAsync(sessionId ?? "", extensionId ?? "", format ?? "", options),
            "unbind" => await UnbindAsync(sessionId ?? "", extensionId),
            "set_format" => await SetFormatAsync(sessionId ?? "", extensionId ?? "", format ?? "", options),
            "status" => GetStatus(extensionId ?? ""),
            "bindings" => GetBindings(sessionId ?? ""),
            "command" => await SendCommandAsync(sessionId ?? "", extensionId ?? "", commandType ?? "", payload),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
        return ResultHelper.FinalizeResult((dynamic)result, null, sessionId);
    }

    /// <summary>
    ///     Lists all registered extensions and their current status.
    /// </summary>
    /// <returns>Result containing the list of extensions with their information.</returns>
    private ListExtensionsResult ListExtensions()
    {
        var extensions = _extensionManager.ListExtensions().ToList();
        var statuses = _extensionManager.GetExtensionStatuses();

        var dtos = extensions.Select(ext =>
        {
            var status = statuses.GetValueOrDefault(ext.Id);
            return new ExtensionInfoDto
            {
                Id = ext.Id,
                Name = ext.DisplayName,
                Version = ext.DisplayVersion,
                Title = ext.DisplayTitle,
                Description = ext.DisplayDescription,
                Author = ext.DisplayAuthor,
                WebsiteUrl = ext.DisplayWebsiteUrl,
                IsAvailable = ext.IsAvailable,
                IsInitializing = status?.IsInitializing ?? false,
                UnavailableReason = ext.UnavailableReason,
                SupportedDocumentTypes = ext.SupportedDocumentTypes,
                InputFormats = ext.InputFormats,
                State = status?.State.ToString()
            };
        }).ToList();

        return new ListExtensionsResult
        {
            Success = true,
            Count = dtos.Count,
            Extensions = dtos
        };
    }

    /// <summary>
    ///     Binds a session to an extension for receiving document snapshots.
    /// </summary>
    /// <param name="sessionId">The session identifier to bind.</param>
    /// <param name="extensionId">The extension identifier to bind to.</param>
    /// <param name="format">The output format for snapshots.</param>
    /// <param name="options">Conversion options for the binding.</param>
    /// <returns>Result containing the binding information or error.</returns>
    private async Task<BindExtensionResult> BindAsync(string sessionId, string extensionId, string format,
        ConversionOptions options)
    {
        if (string.IsNullOrEmpty(sessionId))
            return new BindExtensionResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "sessionId is required for bind operation"
            };

        if (string.IsNullOrEmpty(extensionId))
            return new BindExtensionResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "extensionId is required for bind operation"
            };

        if (string.IsNullOrEmpty(format))
            return new BindExtensionResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "format is required for bind operation"
            };

        var identity = _identityAccessor.GetCurrentIdentity();
        var result = await _sessionBridge.BindAsync(sessionId, extensionId, format, options, identity);

        if (!result.IsSuccess || result.Binding == null)
            return new BindExtensionResult
            {
                Success = false,
                ErrorCode = result.ErrorCode != ExtensionErrorCode.None
                    ? result.ErrorCode
                    : ExtensionErrorCode.InternalError,
                Error = result.Error ?? "Binding operation returned null binding"
            };

        return new BindExtensionResult
        {
            Success = true,
            Binding = new BindingInfoDto
            {
                SessionId = result.Binding.SessionId,
                ExtensionId = result.Binding.ExtensionId,
                OutputFormat = result.Binding.OutputFormat,
                CreatedAt = result.Binding.CreatedAt,
                LastSentAt = result.Binding.LastSentAt
            }
        };
    }

    /// <summary>
    ///     Unbinds a session from an extension or all extensions.
    /// </summary>
    /// <param name="sessionId">The session identifier to unbind.</param>
    /// <param name="extensionId">
    ///     The specific extension identifier to unbind from, or <c>null</c> to unbind from all extensions.
    /// </param>
    /// <returns>Result containing the number of bindings removed.</returns>
    private async Task<UnbindExtensionResult> UnbindAsync(string sessionId, string? extensionId)
    {
        if (string.IsNullOrEmpty(sessionId))
            return new UnbindExtensionResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "sessionId is required for unbind operation",
                SessionId = sessionId,
                ExtensionId = extensionId,
                UnboundCount = 0
            };

        int count;
        if (string.IsNullOrEmpty(extensionId))
            count = await _sessionBridge.UnbindAllAndNotifyAsync(sessionId);
        else
            count = await _sessionBridge.UnbindAndNotifyAsync(sessionId, extensionId) ? 1 : 0;

        return new UnbindExtensionResult
        {
            Success = true,
            SessionId = sessionId,
            ExtensionId = extensionId,
            UnboundCount = count
        };
    }

    /// <summary>
    ///     Changes the output format for an existing session-extension binding.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <param name="extensionId">The extension identifier.</param>
    /// <param name="format">The new output format to set.</param>
    /// <param name="options">Conversion options for the binding.</param>
    /// <returns>Result indicating success or failure of the format change.</returns>
    private async Task<SetFormatResult> SetFormatAsync(string sessionId, string extensionId, string format,
        ConversionOptions options)
    {
        if (string.IsNullOrEmpty(sessionId))
            return new SetFormatResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "sessionId is required for set_format operation",
                SessionId = sessionId,
                ExtensionId = extensionId,
                NewFormat = format
            };

        if (string.IsNullOrEmpty(extensionId))
            return new SetFormatResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "extensionId is required for set_format operation",
                SessionId = sessionId,
                ExtensionId = extensionId,
                NewFormat = format
            };

        if (string.IsNullOrEmpty(format))
            return new SetFormatResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "format is required for set_format operation",
                SessionId = sessionId,
                ExtensionId = extensionId,
                NewFormat = format
            };

        var identity = _identityAccessor.GetCurrentIdentity();
        var result = await _sessionBridge.SetFormatAsync(sessionId, extensionId, format, options, identity);

        return new SetFormatResult
        {
            Success = result.IsSuccess,
            ErrorCode = result.ErrorCode,
            Error = result.Error,
            SessionId = sessionId,
            ExtensionId = extensionId,
            NewFormat = format
        };
    }

    /// <summary>
    ///     Gets the detailed status of a specific extension.
    /// </summary>
    /// <param name="extensionId">The extension identifier to query.</param>
    /// <returns>Result containing the extension status information.</returns>
    private ExtensionStatusResult GetStatus(string extensionId)
    {
        if (string.IsNullOrEmpty(extensionId))
            return new ExtensionStatusResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "extensionId is required for status operation",
                ExtensionId = extensionId,
                Name = "",
                State = "Unknown"
            };

        var statuses = _extensionManager.GetExtensionStatuses();

        if (!statuses.TryGetValue(extensionId, out var status))
            return new ExtensionStatusResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.ExtensionNotFound,
                Error = $"Extension '{extensionId}' not found",
                ExtensionId = extensionId,
                Name = "",
                State = "NotFound"
            };

        var bindings = _sessionBridge.GetBindingsByExtension(extensionId).ToList();

        return new ExtensionStatusResult
        {
            Success = true,
            ExtensionId = status.Id,
            Name = status.Name,
            IsAvailable = status.IsAvailable,
            IsInitializing = status.IsInitializing,
            UnavailableReason = status.UnavailableReason,
            State = status.State.ToString(),
            LastActivity = status.LastActivity,
            RestartCount = status.RestartCount,
            ActiveBindings = bindings.Count
        };
    }

    /// <summary>
    ///     Gets all extension bindings for a specific session.
    /// </summary>
    /// <param name="sessionId">The session identifier to query.</param>
    /// <returns>Result containing the list of bindings for the session.</returns>
    private ExtensionBindingsResult GetBindings(string sessionId)
    {
        if (string.IsNullOrEmpty(sessionId))
            return new ExtensionBindingsResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "sessionId is required for bindings operation",
                SessionId = sessionId,
                Count = 0,
                Bindings = []
            };

        var bindings = _sessionBridge.GetBindings(sessionId)
            .Select(b => new BindingInfoDto
            {
                SessionId = b.SessionId,
                ExtensionId = b.ExtensionId,
                OutputFormat = b.OutputFormat,
                CreatedAt = b.CreatedAt,
                LastSentAt = b.LastSentAt
            })
            .ToList();

        return new ExtensionBindingsResult
        {
            Success = true,
            SessionId = sessionId,
            Count = bindings.Count,
            Bindings = bindings
        };
    }

    /// <summary>
    ///     Sends a command to an extension.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <param name="extensionId">The extension identifier.</param>
    /// <param name="commandType">The type of command to send.</param>
    /// <param name="payload">Optional JSON payload for the command.</param>
    /// <returns>Result containing the command response or error.</returns>
    private async Task<SendCommandResult> SendCommandAsync(
        string sessionId,
        string extensionId,
        string commandType,
        string? payload)
    {
        if (string.IsNullOrEmpty(sessionId))
            return new SendCommandResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "sessionId is required for command operation",
                SessionId = sessionId,
                ExtensionId = extensionId,
                CommandType = commandType
            };

        if (string.IsNullOrEmpty(extensionId))
            return new SendCommandResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "extensionId is required for command operation",
                SessionId = sessionId,
                ExtensionId = extensionId,
                CommandType = commandType
            };

        if (string.IsNullOrEmpty(commandType))
            return new SendCommandResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InvalidParameter,
                Error = "commandType is required for command operation",
                SessionId = sessionId,
                ExtensionId = extensionId,
                CommandType = commandType
            };

        // Parse payload if provided
        Dictionary<string, object>? payloadDict = null;
        if (!string.IsNullOrEmpty(payload))
            try
            {
                payloadDict = JsonSerializer.Deserialize<Dictionary<string, object>>(payload);
            }
            catch (JsonException ex)
            {
                return new SendCommandResult
                {
                    Success = false,
                    ErrorCode = ExtensionErrorCode.InvalidPayload,
                    Error = $"Invalid JSON payload: {ex.Message}",
                    SessionId = sessionId,
                    ExtensionId = extensionId,
                    CommandType = commandType
                };
            }

        // Get extension instance
        var extension = await _extensionManager.GetExtensionAsync(extensionId);
        if (extension == null)
            return new SendCommandResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.ExtensionNotFound,
                Error = $"Extension '{extensionId}' not found",
                SessionId = sessionId,
                ExtensionId = extensionId,
                CommandType = commandType
            };

        // Send command
        try
        {
            var response = await extension.SendCommandAsync(
                sessionId,
                commandType,
                payloadDict);

            if (!response.IsSuccess)
                return new SendCommandResult
                {
                    Success = false,
                    ErrorCode = ExtensionErrorCode.CommandFailed,
                    Error = response.Error ?? "Command execution failed",
                    SessionId = sessionId,
                    ExtensionId = extensionId,
                    CommandId = response.CommandId,
                    CommandType = commandType
                };

            return new SendCommandResult
            {
                Success = true,
                SessionId = sessionId,
                ExtensionId = extensionId,
                CommandId = response.CommandId,
                CommandType = commandType,
                Result = response.Result
            };
        }
        catch (OperationCanceledException)
        {
            return new SendCommandResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.CommandTimeout,
                Error = "Command execution timed out",
                SessionId = sessionId,
                ExtensionId = extensionId,
                CommandType = commandType
            };
        }
        catch (Exception ex)
        {
            return new SendCommandResult
            {
                Success = false,
                ErrorCode = ExtensionErrorCode.InternalError,
                Error = $"Command execution error: {ex.Message}",
                SessionId = sessionId,
                ExtensionId = extensionId,
                CommandType = commandType
            };
        }
    }
}
