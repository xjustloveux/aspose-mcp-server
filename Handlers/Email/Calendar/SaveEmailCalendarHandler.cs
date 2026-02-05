using Aspose.Email;
using Aspose.Email.Calendar;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Calendar;

/// <summary>
///     Handler for saving a calendar appointment to a different format (ICS or MSG).
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SaveEmailCalendarHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "save";

    /// <summary>
    ///     Loads a calendar appointment and saves it to the specified output path and format.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (input ICS file path), outputPath (output file path).
    ///     Optional: format ("ics" or "msg", default: "ics").
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming the appointment was saved.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing, invalid, or format is unsupported.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input ICS file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");
        var format = parameters.GetOptional("format", "ics").ToLowerInvariant();

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Calendar file not found: {path}");

        var appointment = Appointment.Load(path);

        switch (format)
        {
            case "ics":
                appointment.Save(outputPath, AppointmentSaveFormat.Ics);
                break;
            case "msg":
                var message = new MailMessage();
                message.AlternateViews.Add(appointment.RequestApointment());
                message.Save(outputPath, SaveOptions.DefaultMsgUnicode);
                break;
            default:
                throw new ArgumentException(
                    $"Unsupported calendar format: {format}. Supported formats: ics, msg.");
        }

        return new SuccessResult
        {
            Message =
                $"Appointment '{appointment.Summary}' saved to '{outputPath}' in {format.ToUpperInvariant()} format."
        };
    }
}
