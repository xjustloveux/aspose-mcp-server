using Aspose.Email.Calendar;
using Aspose.Email.Calendar.Recurrences;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Calendar;

/// <summary>
///     Handler for setting a recurrence pattern on a calendar appointment.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetRecurrenceEmailCalendarHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "set_recurrence";

    /// <summary>
    ///     Sets a recurrence pattern on the specified appointment and saves the result.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (input ICS file path), outputPath (output ICS file path),
    ///     pattern ("daily", "weekly", "monthly", "yearly").
    ///     Optional: interval (default: 1), count (default: 10).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming the recurrence pattern was set.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing, invalid, or pattern is unsupported.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input ICS file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");
        var pattern = parameters.GetRequired<string>("pattern").ToLowerInvariant();
        var interval = parameters.GetOptional("interval", 1);
        var count = parameters.GetOptional("count", 10);

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Calendar file not found: {path}");

        var appointment = Appointment.Load(path);

        var recurrence = CreateRecurrencePattern(pattern, interval, count);
        appointment.Recurrence = recurrence;
        appointment.Save(outputPath, AppointmentSaveFormat.Ics);

        return new SuccessResult
        {
            Message = $"Recurrence pattern '{pattern}' set on appointment '{appointment.Summary}' " +
                      $"(interval: {interval}, count: {count}). Saved to '{outputPath}'."
        };
    }

    /// <summary>
    ///     Creates a recurrence pattern based on the specified pattern name, interval, and count.
    /// </summary>
    /// <param name="pattern">The recurrence pattern name (daily, weekly, monthly, yearly).</param>
    /// <param name="interval">The recurrence interval.</param>
    /// <param name="count">The number of occurrences.</param>
    /// <returns>A configured <see cref="RecurrencePattern" />.</returns>
    /// <exception cref="ArgumentException">Thrown when the pattern is not recognized.</exception>
    private static RecurrencePattern CreateRecurrencePattern(string pattern, int interval, int count)
    {
        return pattern switch
        {
            "daily" => new DailyRecurrencePattern(count, interval),
            "weekly" => new WeeklyRecurrencePattern(count, interval),
            "monthly" => new MonthlyRecurrencePattern(count, interval),
            "yearly" => new YearlyRecurrencePattern { Interval = interval, Occurs = count },
            _ => throw new ArgumentException(
                $"Unknown recurrence pattern: {pattern}. Supported patterns: daily, weekly, monthly, yearly.")
        };
    }
}
