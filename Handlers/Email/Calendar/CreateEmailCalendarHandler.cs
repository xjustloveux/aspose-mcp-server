using Aspose.Email;
using Aspose.Email.Calendar;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Calendar;

/// <summary>
///     Handler for creating a new calendar appointment and saving it as an ICS file.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class CreateEmailCalendarHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "create";

    /// <summary>
    ///     Creates a new calendar appointment and saves it to the specified output path.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: outputPath (ICS file path).
    ///     Optional: summary, description, startDate (ISO 8601), endDate (ISO 8601), location, attendees (comma-separated).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming the appointment was created.</returns>
    /// <exception cref="ArgumentException">Thrown when the outputPath parameter is missing or invalid.</exception>
    /// <exception cref="FormatException">Thrown when startDate or endDate cannot be parsed.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var outputPath = parameters.GetRequired<string>("outputPath");
        var summary = parameters.GetOptional<string?>("summary");
        var description = parameters.GetOptional<string?>("description");
        var startDate = parameters.GetOptional<string?>("startDate");
        var endDate = parameters.GetOptional<string?>("endDate");
        var location = parameters.GetOptional<string?>("location");
        var attendees = parameters.GetOptional<string?>("attendees");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var start = string.IsNullOrEmpty(startDate) ? DateTime.Now : DateTime.Parse(startDate);
        var end = string.IsNullOrEmpty(endDate) ? start.AddHours(1) : DateTime.Parse(endDate);

        var appointment = new Appointment(
            location ?? "",
            summary ?? "New Appointment",
            description ?? "",
            start,
            end,
            new MailAddress("organizer@example.com"),
            new MailAddressCollection());

        if (!string.IsNullOrEmpty(attendees))
            foreach (var attendee in attendees.Split(',',
                         StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                appointment.Attendees.Add(new MailAddress(attendee));

        appointment.Save(outputPath, AppointmentSaveFormat.Ics);

        return new SuccessResult
        {
            Message = $"Calendar appointment '{appointment.Summary}' created and saved to '{outputPath}'."
        };
    }
}
