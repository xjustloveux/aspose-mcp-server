using System.ComponentModel;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Email;

/// <summary>
///     Tool for managing calendar appointments (create, get_info, save, set_recurrence, set_attendees).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Email.Calendar")]
[McpServerToolType]
public class EmailCalendarTool
{
    /// <summary>
    ///     Handler registry for email calendar operations.
    /// </summary>
    private readonly HandlerRegistry<object> _handlerRegistry;

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailCalendarTool" /> class.
    /// </summary>
    public EmailCalendarTool()
    {
        _handlerRegistry =
            HandlerRegistry<object>.CreateFromNamespace("AsposeMcpServer.Handlers.Email.Calendar");
    }

    /// <summary>
    ///     Executes a calendar operation (create, get_info, save, set_recurrence, set_attendees).
    /// </summary>
    /// <param name="operation">The operation to perform: create, get_info, save, set_recurrence, set_attendees.</param>
    /// <param name="path">Input calendar file path (.ics) (required for get_info, save, set_recurrence, set_attendees).</param>
    /// <param name="outputPath">Output file path (required for create, save, set_recurrence, set_attendees).</param>
    /// <param name="summary">Appointment summary/title (optional for create).</param>
    /// <param name="description">Appointment description (optional for create).</param>
    /// <param name="startDate">Start date in ISO 8601 format (optional for create).</param>
    /// <param name="endDate">End date in ISO 8601 format (optional for create).</param>
    /// <param name="location">Appointment location (optional for create).</param>
    /// <param name="attendees">Comma-separated attendee email addresses (optional for create, required for set_attendees).</param>
    /// <param name="format">Output format: "ics" or "msg" (optional for save, default: "ics").</param>
    /// <param name="pattern">Recurrence pattern: "daily", "weekly", "monthly", "yearly" (required for set_recurrence).</param>
    /// <param name="interval">Recurrence interval (optional for set_recurrence, default: 1).</param>
    /// <param name="count">Number of recurrences (optional for set_recurrence, default: 10).</param>
    /// <returns>Operation result depending on the operation type.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "email_calendar",
        Title = "Email Calendar Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage calendar appointments (ICS files). Supports 5 operations: create, get_info, save, set_recurrence, set_attendees.

Usage examples:
- Create appointment: email_calendar(operation='create', outputPath='meeting.ics', summary='Team Meeting', startDate='2024-01-15T10:00:00', endDate='2024-01-15T11:00:00')
- Get info: email_calendar(operation='get_info', path='meeting.ics')
- Save as MSG: email_calendar(operation='save', path='meeting.ics', outputPath='meeting.msg', format='msg')
- Set recurrence: email_calendar(operation='set_recurrence', path='meeting.ics', outputPath='recurring.ics', pattern='weekly', interval=1, count=52)
- Set attendees: email_calendar(operation='set_attendees', path='meeting.ics', outputPath='updated.ics', attendees='user1@example.com,user2@example.com')

Supported formats: ICS (iCalendar), MSG (Outlook)")]
    public object Execute(
        [Description(@"Operation to perform.
- 'create': Create a new appointment (required params: outputPath; optional: summary, description, startDate, endDate, location, attendees)
- 'get_info': Get appointment information (required params: path)
- 'save': Save appointment to a different format (required params: path, outputPath; optional: format)
- 'set_recurrence': Set recurrence pattern (required params: path, outputPath, pattern; optional: interval, count)
- 'set_attendees': Set attendees (required params: path, outputPath, attendees)")]
        string operation,
        [Description("Input calendar file path (.ics) (required for get_info, save, set_recurrence, set_attendees)")]
        string? path = null,
        [Description("Output file path (required for create, save, set_recurrence, set_attendees)")]
        string? outputPath = null,
        [Description("Appointment summary/title (optional for create)")]
        string? summary = null,
        [Description("Appointment description (optional for create)")]
        string? description = null,
        [Description("Start date in ISO 8601 format, e.g. '2024-01-15T10:00:00' (optional for create)")]
        string? startDate = null,
        [Description("End date in ISO 8601 format, e.g. '2024-01-15T11:00:00' (optional for create)")]
        string? endDate = null,
        [Description("Appointment location (optional for create)")]
        string? location = null,
        [Description("Comma-separated attendee email addresses (optional for create, required for set_attendees)")]
        string? attendees = null,
        [Description("Output format: 'ics' or 'msg' (optional for save, default: 'ics')")]
        string? format = null,
        [Description("Recurrence pattern: 'daily', 'weekly', 'monthly', 'yearly' (required for set_recurrence)")]
        string? pattern = null,
        [Description("Recurrence interval (optional for set_recurrence, default: 1)")]
        int? interval = null,
        [Description("Number of recurrences (optional for set_recurrence, default: 10)")]
        int? count = null)
    {
        var parameters = BuildParameters(path, outputPath, summary, description, startDate, endDate, location,
            attendees, format, pattern, interval, count);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<object>
        {
            Document = new object(),
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        var effectiveOutputPath = string.Equals(operation, "get_info", StringComparison.OrdinalIgnoreCase)
            ? path
            : outputPath;

        return ResultHelper.FinalizeResult((dynamic)result, effectiveOutputPath, (string?)null);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="path">The input calendar file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="summary">The appointment summary.</param>
    /// <param name="description">The appointment description.</param>
    /// <param name="startDate">The start date string.</param>
    /// <param name="endDate">The end date string.</param>
    /// <param name="location">The appointment location.</param>
    /// <param name="attendees">The comma-separated attendee list.</param>
    /// <param name="format">The output format.</param>
    /// <param name="pattern">The recurrence pattern.</param>
    /// <param name="interval">The recurrence interval.</param>
    /// <param name="count">The recurrence count.</param>
    /// <returns>OperationParameters configured for the calendar operation.</returns>
    private static OperationParameters BuildParameters(
        string? path,
        string? outputPath,
        string? summary,
        string? description,
        string? startDate,
        string? endDate,
        string? location,
        string? attendees,
        string? format,
        string? pattern,
        int? interval,
        int? count)
    {
        var parameters = new OperationParameters();
        parameters.SetIfNotNull("path", path);
        parameters.SetIfNotNull("outputPath", outputPath);
        parameters.SetIfNotNull("summary", summary);
        parameters.SetIfNotNull("description", description);
        parameters.SetIfNotNull("startDate", startDate);
        parameters.SetIfNotNull("endDate", endDate);
        parameters.SetIfNotNull("location", location);
        parameters.SetIfNotNull("attendees", attendees);
        parameters.SetIfNotNull("format", format);
        parameters.SetIfNotNull("pattern", pattern);
        parameters.SetIfHasValue("interval", interval);
        parameters.SetIfHasValue("count", count);
        return parameters;
    }
}
