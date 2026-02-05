using Aspose.Email;
using Aspose.Email.Calendar;
using AsposeMcpServer.Handlers.Email.Calendar;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Calendar;

/// <summary>
///     Tests for <see cref="SaveEmailCalendarHandler" />.
/// </summary>
public class SaveEmailCalendarHandlerTests : HandlerTestBase<object>
{
    private readonly SaveEmailCalendarHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Save()
    {
        Assert.Equal("save", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Creates a test ICS file with the specified properties.
    /// </summary>
    private string CreateTestIcsFile(string fileName, string summary, string description = "",
        string location = "", params string[] attendees)
    {
        var filePath = Path.Combine(TestDir, fileName);
        var attendeeCollection = new MailAddressCollection();
        foreach (var attendee in attendees)
            attendeeCollection.Add(new MailAddress(attendee));

        var appointment = new Appointment(
            location,
            summary,
            description,
            DateTime.Now,
            DateTime.Now.AddHours(1),
            new MailAddress("organizer@example.com"),
            attendeeCollection);
        appointment.Save(filePath, AppointmentSaveFormat.Ics);
        return filePath;
    }

    #endregion

    #region Basic Save Operations

    [Fact]
    public void Execute_SaveAsIcs_CreatesIcsFile()
    {
        var inputPath = CreateTestIcsFile("save_input.ics", "Save Test");
        var outputPath = Path.Combine(TestDir, "save_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "format", "ics" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Save Test", result.Message);
        Assert.Contains("ICS", result.Message);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("Save Test", appointment.Summary);
    }

    [SkippableFact]
    public void Execute_SaveAsMsg_CreatesMsgFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Email, "MSG conversion requires license");

        var inputPath = CreateTestIcsFile("save_msg_input.ics", "MSG Save Test");
        var outputPath = Path.Combine(TestDir, "save_output.msg");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "format", "msg" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("MSG", result.Message);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_WithDefaultFormat_SavesAsIcs()
    {
        var inputPath = CreateTestIcsFile("default_fmt_input.ics", "Default Format Test");
        var outputPath = Path.Combine(TestDir, "default_fmt_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("ICS", result.Message);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("Default Format Test", appointment.Summary);
    }

    [Fact]
    public void Execute_PreservesAppointmentData()
    {
        var inputPath = CreateTestIcsFile("preserve_input.ics", "Preserved Meeting",
            "Important description", "Room 42", "attendee@example.com");
        var outputPath = Path.Combine(TestDir, "preserve_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "format", "ics" }
        });

        _handler.Execute(context, parameters);

        var saved = Appointment.Load(outputPath);
        Assert.Equal("Preserved Meeting", saved.Summary);
        Assert.Equal("Important description", saved.Description);
        Assert.Equal("Room 42", saved.Location);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", Path.Combine(TestDir, "output.ics") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var inputPath = CreateTestIcsFile("missing_output.ics", "Test");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath }
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
            { "outputPath", Path.Combine(TestDir, "output.ics") }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedFormat_ThrowsArgumentException()
    {
        var inputPath = CreateTestIcsFile("unsupported_fmt.ics", "Test");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", Path.Combine(TestDir, "output.xyz") },
            { "format", "xyz" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
