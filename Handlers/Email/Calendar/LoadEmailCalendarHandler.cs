using Aspose.Email.Calendar;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Email.Calendar;

namespace AsposeMcpServer.Handlers.Email.Calendar;

/// <summary>
///     Handler for loading and reading calendar appointment information from an ICS file.
/// </summary>
[ResultType(typeof(AppointmentEmailInfo))]
public class LoadEmailCalendarHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "get_info";

    /// <summary>
    ///     Loads a calendar appointment from the specified ICS file and returns its information.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (ICS file path).
    /// </param>
    /// <returns>An <see cref="AppointmentEmailInfo" /> containing the appointment details.</returns>
    /// <exception cref="ArgumentException">Thrown when the path parameter is missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the ICS file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Calendar file not found: {path}");

        var appointment = Appointment.Load(path);

        var attendeeList = new List<string>();
        foreach (var attendee in appointment.Attendees)
            attendeeList.Add(attendee.Address);

        return new AppointmentEmailInfo
        {
            Summary = appointment.Summary,
            Description = appointment.Description,
            Location = appointment.Location,
            StartDate = appointment.StartDate.ToString("o"),
            EndDate = appointment.EndDate.ToString("o"),
            Attendees = attendeeList.Count > 0 ? attendeeList : null,
            IsRecurring = appointment.Recurrence != null,
            Message = $"Loaded appointment '{appointment.Summary}' from '{path}'."
        };
    }
}
