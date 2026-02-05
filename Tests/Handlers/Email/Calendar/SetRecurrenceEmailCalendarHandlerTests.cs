using Aspose.Email;
using Aspose.Email.Calendar;
using AsposeMcpServer.Handlers.Email.Calendar;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Email.Calendar;

/// <summary>
///     Tests for <see cref="SetRecurrenceEmailCalendarHandler" />.
/// </summary>
public class SetRecurrenceEmailCalendarHandlerTests : HandlerTestBase<object>
{
    private readonly SetRecurrenceEmailCalendarHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetRecurrence()
    {
        Assert.Equal("set_recurrence", _handler.Operation);
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

    #endregion

    #region Basic Recurrence Operations

    [Theory]
    [InlineData("daily")]
    [InlineData("weekly")]
    [InlineData("monthly")]
    [InlineData("yearly")]
    public void Execute_WithValidPattern_SetsRecurrence(string pattern)
    {
        var inputPath = CreateTestIcsFile("recurrence_input.ics", "Recurring Meeting");
        var outputPath = Path.Combine(TestDir, $"recurrence_{pattern}_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "pattern", pattern }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains(pattern, result.Message);
        Assert.Contains("Recurring Meeting", result.Message);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Execute_WithCustomInterval_SetsInterval()
    {
        var inputPath = CreateTestIcsFile("interval_input.ics", "Bi-weekly Meeting");
        var outputPath = Path.Combine(TestDir, "interval_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "pattern", "weekly" },
            { "interval", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("interval: 2", result.Message);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.NotNull(appointment.Recurrence);
    }

    [Fact]
    public void Execute_WithCustomCount_SetsCount()
    {
        var inputPath = CreateTestIcsFile("count_input.ics", "Limited Meeting");
        var outputPath = Path.Combine(TestDir, "count_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "pattern", "daily" },
            { "count", 5 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("count: 5", result.Message);
        Assert.True(File.Exists(outputPath));

        var appointment = Appointment.Load(outputPath);
        Assert.NotNull(appointment.Recurrence);
    }

    [Fact]
    public void Execute_WithDefaultIntervalAndCount_UsesDefaults()
    {
        var inputPath = CreateTestIcsFile("defaults_input.ics", "Default Recurrence");
        var outputPath = Path.Combine(TestDir, "defaults_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "pattern", "daily" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("interval: 1", result.Message);
        Assert.Contains("count: 10", result.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_PreservesAppointmentData()
    {
        var inputPath = CreateTestIcsFile("preserve_recur_input.ics", "Preserve Test",
            "Keep this description", "Room 101");
        var outputPath = Path.Combine(TestDir, "preserve_recur_output.ics");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", outputPath },
            { "pattern", "weekly" }
        });

        _handler.Execute(context, parameters);

        var appointment = Appointment.Load(outputPath);
        Assert.Equal("Preserve Test", appointment.Summary);
        Assert.Equal("Keep this description", appointment.Description);
        Assert.Equal("Room 101", appointment.Location);
        Assert.NotNull(appointment.Recurrence);
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
            { "pattern", "daily" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var inputPath = CreateTestIcsFile("no_output.ics", "Test");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "pattern", "daily" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutPattern_ThrowsArgumentException()
    {
        var inputPath = CreateTestIcsFile("no_pattern.ics", "Test");
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
            { "pattern", "daily" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedPattern_ThrowsArgumentException()
    {
        var inputPath = CreateTestIcsFile("bad_pattern.ics", "Test");
        var context = CreateContext(new object());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", inputPath },
            { "outputPath", Path.Combine(TestDir, "output.ics") },
            { "pattern", "biweekly" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
