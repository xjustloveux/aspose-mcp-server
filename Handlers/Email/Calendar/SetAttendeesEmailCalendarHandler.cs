using Aspose.Email;
using Aspose.Email.Calendar;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Calendar;

/// <summary>
///     Handler for setting attendees on a calendar appointment.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetAttendeesEmailCalendarHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "set_attendees";

    /// <summary>
    ///     Sets the attendees on the specified appointment, replacing existing attendees, and saves the result.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (input ICS file path), outputPath (output ICS file path),
    ///     attendees (comma-separated email addresses).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming the attendees were set.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input ICS file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");
        var attendees = parameters.GetRequired<string>("attendees");

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Calendar file not found: {path}");

        var appointment = Appointment.Load(path);

        appointment.Attendees.Clear();
        var attendeeList = attendees.Split(',',
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var attendee in attendeeList)
            appointment.Attendees.Add(new MailAddress(attendee));

        appointment.Save(outputPath, AppointmentSaveFormat.Ics);

        return new SuccessResult
        {
            Message = $"Set {attendeeList.Length} attendee(s) on appointment '{appointment.Summary}'. " +
                      $"Saved to '{outputPath}'."
        };
    }
}
