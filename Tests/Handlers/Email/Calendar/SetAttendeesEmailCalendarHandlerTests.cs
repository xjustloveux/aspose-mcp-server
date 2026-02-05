using Aspose.Email;
using Aspose.Email.Calendar;
using AsposeMcpServer.Handlers.Email.Calendar;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Calendar;

/// <summary>
///     Tests for <see cref="SetAttendeesEmailCalendarHandler" />.
/// </summary>
public class SetAttendeesEmailCalendarHandlerTests : HandlerTestBase<object>
{
    private readonly SetAttendeesEmailCalendarHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetAttendees()
    {
        Assert.Equal("set_attendees", _handler.Operation);
    }

    #endregion

    #region Basic Set Attendees Operations

    [Fact]
    public void Execute_WithSingleAttendee_SetsOneAttendee()
    {
        var inputPath = CreateTestIcsFile("single_attendee_input.ics", "Team Meeting");
        var outputPath = Path.Combine(TestDir, "single_attendee_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "attendees", "alice@example.com" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("1 attendee(s)", result.Message);
        Assert.Contains("Team Meeting", result.Message);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Single(appointment.Attendees);
    }

    [Fact]
    public void Execute_WithMultipleAttendees_SetsAllAttendees()
    {
        var inputPath = CreateTestIcsFile("multi_attendee_input.ics", "Group Meeting");
        var outputPath = Path.Combine(TestDir, "multi_attendee_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "attendees", "alice@example.com,bob@example.com,charlie@example.com" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("3 attendee(s)", result.Message);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Equal(3, appointment.Attendees.Count);
    }

    [Fact]
    public void Execute_WithSpaceSeparatedAttendees_TrimsWhitespace()
    {
        var inputPath = CreateTestIcsFile("trim_input.ics", "Trim Test");
        var outputPath = Path.Combine(TestDir, "trim_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "attendees", " alice@example.com , bob@example.com " }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("2 attendee(s)", result.Message);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Equal(2, appointment.Attendees.Count);
    }

    [Fact]
    public void Execute_ReplacesExistingAttendees()
    {
        var inputPath = CreateTestIcsFileWithAttendees("replace_input.ics", "Replace Test",
            "old1@example.com", "old2@example.com");
        var outputPath = Path.Combine(TestDir, "replace_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "attendees", "new1@example.com,new2@example.com,new3@example.com" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("3 attendee(s)", result.Message);

        var appointment = Appointment.Load(outputPath);
        Assert.Equal(3, appointment.Attendees.Count);
    }

    [Fact]
    public void Execute_PreservesAppointmentData()
    {
        var inputPath = CreateTestIcsFile("preserve_att_input.ics", "Important Meeting",
            "Keep this", "Room 42");
        var outputPath = Path.Combine(TestDir, "preserve_att_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "attendees", "attendee@example.com" }
        });

        _handler.Execute(context, parameters);

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("Important Meeting", appointment.Summary);
        Assert.Equal("Keep this", appointment.Description);
        Assert.Equal("Room 42", appointment.Location);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", Path.Combine(TestDir, "output.ics") },
            { "attendees", "test@example.com" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var inputPath = CreateTestIcsFile("no_output_att.ics", "Test");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "attendees", "test@example.com" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutAttendees_ThrowsArgumentException()
    {
        var inputPath = CreateTestIcsFile("no_attendees_param.ics", "Test");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", Path.Combine(TestDir, "output.ics") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.ics") },
            { "outputPath", Path.Combine(TestDir, "output.ics") },
            { "attendees", "test@example.com" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test ICS file with the specified properties.
    /// </summary>
    private string CreateTestIcsFile(string fileName, string summary, string description = "",
        string location = "")
    {
        var filePath = Path.Combine(TestDir, fileName);
        var appointment = new Appointment(
            location,
            summary,
            description,
            DateTime.Now,
            DateTime.Now.AddHours(1),
            new MailAddress("organizer@example.com"),
            new MailAddressCollection());
        appointment.Save(filePath, AppointmentSaveFormat.Ics);
        return filePath;
    }

    /// <summary>
    ///     Creates a test ICS file with attendees already set.
    /// </summary>
    private string CreateTestIcsFileWithAttendees(string fileName, string summary, params string[] attendees)
    {
        var filePath = Path.Combine(TestDir, fileName);
        var attendeeCollection = new MailAddressCollection();
        foreach (var attendee in attendees)
            attendeeCollection.Add(new MailAddress(attendee));

        var appointment = new Appointment(
            "",
            summary,
            "",
            DateTime.Now,
            DateTime.Now.AddHours(1),
            new MailAddress("organizer@example.com"),
            attendeeCollection);
        appointment.Save(filePath, AppointmentSaveFormat.Ics);
        return filePath;
    }

    #endregion
}
