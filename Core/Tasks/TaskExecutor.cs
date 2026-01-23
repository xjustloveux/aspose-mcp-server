using System.Text.Json;
using AsposeMcpServer.Tools.Conversion;

namespace AsposeMcpServer.Core.Tasks;

/// <summary>
///     Executes tasks asynchronously with proper error handling.
/// </summary>
public sealed class TaskExecutor
{
    /// <summary>
    ///     Tools that support async execution.
    /// </summary>
    public static readonly HashSet<string> SupportedTools = new(StringComparer.OrdinalIgnoreCase)
    {
        "convert_to_pdf",
        "convert_document"
    };

    private readonly ILogger<TaskExecutor>? _logger;
    private readonly IServiceProvider _services;
    private readonly TaskStore _store;

    /// <summary>
    ///     Creates a new task executor.
    /// </summary>
    /// <param name="store">Task store.</param>
    /// <param name="services">Service provider for tool resolution.</param>
    /// <param name="logger">Optional logger.</param>
    /// <exception cref="ArgumentNullException">Thrown when store or services is null.</exception>
    public TaskExecutor(
        TaskStore store,
        IServiceProvider services,
        ILogger<TaskExecutor>? logger = null)
    {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _services = services ?? throw new ArgumentNullException(nameof(services));
        _logger = logger;
    }

    /// <summary>
    ///     Checks if a tool supports async execution.
    /// </summary>
    /// <param name="toolName">The tool name to check.</param>
    /// <returns>True if the tool supports async execution.</returns>
    public static bool SupportsAsync(string toolName)
    {
        return SupportedTools.Contains(toolName);
    }

    /// <summary>
    ///     Executes a task asynchronously.
    /// </summary>
    /// <param name="taskId">The task ID to execute.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    public async Task ExecuteAsync(string taskId)
    {
        var task = _store.GetTask(taskId);
        if (task == null)
        {
            _logger?.LogWarning("Task {TaskId} not found for execution", taskId);
            return;
        }

        var cancellationToken = task.CancellationTokenSource.Token;

        _store.UpdateTaskStatus(taskId, TaskStatus.Working, "Processing...");

        try
        {
            _logger?.LogInformation(
                "Executing task {TaskId} for tool {ToolName}",
                taskId, task.ToolName);

            var result = await ExecuteToolAsync(
                task.ToolName,
                task.Arguments,
                cancellationToken);

            _store.UpdateTaskStatus(
                taskId,
                TaskStatus.Completed,
                "Task completed successfully",
                result);

            _logger?.LogInformation("Task {TaskId} completed successfully", taskId);
        }
        catch (OperationCanceledException)
        {
            _store.UpdateTaskStatus(
                taskId,
                TaskStatus.Cancelled,
                "Task was cancelled");

            _logger?.LogInformation("Task {TaskId} was cancelled", taskId);
        }
        catch (FileNotFoundException ex)
        {
            _store.UpdateTaskStatus(
                taskId,
                TaskStatus.Failed,
                "File not found",
                errorMessage: ex.Message);

            _logger?.LogWarning(ex, "Task {TaskId} failed: file not found", taskId);
        }
        catch (UnauthorizedAccessException ex)
        {
            _store.UpdateTaskStatus(
                taskId,
                TaskStatus.Failed,
                "Access denied",
                errorMessage: ex.Message);

            _logger?.LogWarning(ex, "Task {TaskId} failed: access denied", taskId);
        }
        catch (ArgumentException ex)
        {
            _store.UpdateTaskStatus(
                taskId,
                TaskStatus.Failed,
                "Invalid argument",
                errorMessage: ex.Message);

            _logger?.LogWarning(ex, "Task {TaskId} failed: invalid argument", taskId);
        }
        catch (NotSupportedException ex)
        {
            _store.UpdateTaskStatus(
                taskId,
                TaskStatus.Failed,
                "Not supported",
                errorMessage: ex.Message);

            _logger?.LogWarning(ex, "Task {TaskId} failed: not supported", taskId);
        }
        catch (Exception ex)
        {
            _store.UpdateTaskStatus(
                taskId,
                TaskStatus.Failed,
                "Task failed",
                errorMessage: ex.Message);

            _logger?.LogError(ex, "Task {TaskId} failed with error", taskId);
        }
    }

    private async Task<string> ExecuteToolAsync(
        string toolName,
        JsonElement arguments,
        CancellationToken cancellationToken)
    {
        return toolName.ToLowerInvariant() switch
        {
            "convert_to_pdf" => await ExecuteConvertToPdfAsync(arguments, cancellationToken),
            "convert_document" => await ExecuteConvertDocumentAsync(arguments, cancellationToken),
            _ => throw new NotSupportedException($"Tool '{toolName}' does not support async execution")
        };
    }

    private async Task<string> ExecuteConvertToPdfAsync(
        JsonElement arguments,
        CancellationToken cancellationToken)
    {
        return await Task.Run(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();

            var tool = _services.GetRequiredService<ConvertToPdfTool>();

            var inputPath = GetOptionalString(arguments, "inputPath");
            var sessionId = GetOptionalString(arguments, "sessionId");
            var outputPath = GetOptionalString(arguments, "outputPath");

            var result = tool.Execute(inputPath, sessionId, outputPath);
            return JsonSerializer.Serialize(result);
        }, cancellationToken);
    }

    private async Task<string> ExecuteConvertDocumentAsync(
        JsonElement arguments,
        CancellationToken cancellationToken)
    {
        return await Task.Run(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();

            var tool = _services.GetRequiredService<ConvertDocumentTool>();

            var inputPath = GetOptionalString(arguments, "inputPath");
            var sessionId = GetOptionalString(arguments, "sessionId");
            var outputPath = GetOptionalString(arguments, "outputPath");

            var result = tool.Execute(inputPath, sessionId, outputPath);
            return JsonSerializer.Serialize(result);
        }, cancellationToken);
    }

    private static string? GetOptionalString(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop) &&
            prop.ValueKind == JsonValueKind.String)
            return prop.GetString();

        return null;
    }
}
