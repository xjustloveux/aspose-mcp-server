using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Manages extension process cleanup to ensure child processes are terminated
///     when the parent process exits unexpectedly.
///     Uses Windows Job Objects on Windows and process tracking on other platforms.
/// </summary>
/// <remarks>
///     <para>
///         This class provides automatic cleanup of extension processes in two scenarios:
///         <list type="number">
///             <item>Normal shutdown via <see cref="Dispose" /></item>
///             <item>Unexpected parent process termination via ProcessExit event and Job Objects</item>
///         </list>
///     </para>
///     <para>
///         On Windows, a Job Object with <c>JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE</c> ensures
///         that all assigned processes are automatically terminated when the parent process exits,
///         even if the parent crashes without cleanup code executing.
///     </para>
///     <para>
///         On non-Windows platforms, the ProcessExit event handler provides best-effort cleanup.
///         This is less reliable than Job Objects but provides some protection.
///     </para>
/// </remarks>
public sealed class ProcessCleanupManager : IDisposable
{
    /// <summary>
    ///     Grace period in seconds after process start before zombie detection kicks in.
    ///     This allows normal process initialization without false positives.
    /// </summary>
    private const int ZombieDetectionGracePeriodSeconds = 5;

    /// <summary>
    ///     Interval in milliseconds to wait between CPU time samples for zombie detection.
    /// </summary>
    private const int CpuSampleIntervalMs = 100;

    /// <summary>
    ///     Logger instance for diagnostic output.
    /// </summary>
    private readonly ILogger? _logger;

    /// <summary>
    ///     Collection of tracked processes for cleanup on non-Windows platforms.
    /// </summary>
    private readonly ConcurrentDictionary<int, Process> _trackedProcesses = new();

    /// <summary>
    ///     Whether this instance has been disposed.
    /// </summary>
    private bool _disposed;

    /// <summary>
    ///     Windows Job Object handle (Windows only).
    /// </summary>
    private SafeHandle? _jobHandle;

    /// <summary>
    ///     Whether the Job Object was successfully created and configured.
    /// </summary>
    private bool _jobObjectValid;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ProcessCleanupManager" /> class.
    /// </summary>
    /// <param name="logger">Optional logger instance.</param>
    public ProcessCleanupManager(ILogger? logger = null)
    {
        _logger = logger;

        if (OperatingSystem.IsWindows())
            InitializeWindowsJobObject();

        AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;

        AppDomain.CurrentDomain.ProcessExit -= OnProcessExit;

        KillTrackedProcesses();

        if (_jobHandle != null)
        {
            _jobHandle.Dispose();
            _jobHandle = null;
        }

        _trackedProcesses.Clear();
    }

    /// <summary>
    ///     Registers a process for automatic cleanup when the parent process exits.
    /// </summary>
    /// <param name="process">The process to register.</param>
    public void RegisterProcess(Process process)
    {
        if (_disposed || process.HasExited)
            return;

        if (OperatingSystem.IsWindows() && _jobObjectValid)
            try
            {
                if (!AssignProcessToJobObject(_jobHandle!, process.SafeHandle))
                {
                    var error = Marshal.GetLastWin32Error();
                    _logger?.LogDebug(
                        "Failed to assign process {Pid} to job object, error: {Error}",
                        process.Id, error);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogDebug(ex,
                    "Failed to assign process {Pid} to job object",
                    process.Id);
            }

        var processId = process.Id;
        _trackedProcesses.TryAdd(processId, process);

        process.Exited += (_, _) => _trackedProcesses.TryRemove(processId, out _);

        if (process.HasExited)
            _trackedProcesses.TryRemove(processId, out _);
    }

    /// <summary>
    ///     Unregisters a process from cleanup tracking.
    /// </summary>
    /// <param name="process">The process to unregister.</param>
    public void UnregisterProcess(Process process)
    {
        _trackedProcesses.TryRemove(process.Id, out _);
    }

    /// <summary>
    ///     Checks if a process is in a zombie or unresponsive state.
    ///     A zombie process has exited but not been reaped, or appears running but is unresponsive.
    /// </summary>
    /// <param name="process">The process to check.</param>
    /// <returns>True if the process appears to be a zombie.</returns>
    /// <remarks>
    ///     <para>Zombie detection heuristics (in order of execution):</para>
    ///     <list type="number">
    ///         <item>If <see cref="Process.HasExited" /> returns true after refresh, process is zombie</item>
    ///         <item>On Windows, if handle access throws <see cref="InvalidOperationException" />, process is zombie</item>
    ///         <item>Grace period: skip checks for processes started within last 5 seconds</item>
    ///         <item>CPU time sample: if CPU time unchanged over 100ms AND thread count is 0, process is zombie</item>
    ///     </list>
    ///     <para>
    ///         This method has some limitations:
    ///         <list type="bullet">
    ///             <item>May return false negatives for I/O-bound processes with zero CPU activity</item>
    ///             <item>Thread.Sleep blocks the calling thread briefly</item>
    ///             <item>Some edge cases may not be detected on non-Windows platforms</item>
    ///         </list>
    ///     </para>
    /// </remarks>
    public bool IsZombieProcess(Process process)
    {
        if (process.HasExited)
            return false;

        try
        {
            process.Refresh();

            if (process.HasExited)
                return true;

            if (OperatingSystem.IsWindows())
                try
                {
                    _ = process.Handle;
                }
                catch (InvalidOperationException)
                {
                    return true;
                }

            var currentTime = DateTime.UtcNow;
            var processTime = process.StartTime.ToUniversalTime();
            if ((currentTime - processTime).TotalSeconds < ZombieDetectionGracePeriodSeconds)
                return false;

            try
            {
                var cpuTime = process.TotalProcessorTime;
                Thread.Sleep(CpuSampleIntervalMs);
                process.Refresh();

                if (process.HasExited)
                    return true;

                var newCpuTime = process.TotalProcessorTime;

                if (cpuTime == newCpuTime)
                    try
                    {
                        var threads = process.Threads;
                        if (threads.Count == 0)
                            return true;
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogTrace(ex, "Failed to enumerate threads for process {Pid}", process.Id);
                    }
            }
            catch (InvalidOperationException)
            {
                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Error checking zombie state for process {Pid}", process.Id);
            return false;
        }
    }

    /// <summary>
    ///     Gets a list of all currently tracked processes that appear to be zombies.
    /// </summary>
    /// <returns>List of zombie process IDs.</returns>
    public IReadOnlyList<int> GetZombieProcesses()
    {
        var zombies = new List<int>();

        foreach (var kvp in _trackedProcesses)
            try
            {
                if (IsZombieProcess(kvp.Value))
                    zombies.Add(kvp.Key);
            }
            catch (Exception ex)
            {
                _logger?.LogTrace(ex, "Failed to check zombie status for process {Pid}", kvp.Key);
            }

        return zombies;
    }

    /// <summary>
    ///     Handles the ProcessExit event to ensure cleanup on normal exit.
    /// </summary>
    private void OnProcessExit(object? sender, EventArgs e)
    {
        KillTrackedProcesses();
    }

    /// <summary>
    ///     Kills all tracked processes that are still running.
    /// </summary>
    private void KillTrackedProcesses()
    {
        foreach (var kvp in _trackedProcesses)
            try
            {
                var process = kvp.Value;
                if (!process.HasExited)
                {
                    process.Kill(true);
                    _logger?.LogDebug("Killed orphan extension process {Pid}", kvp.Key);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogDebug(ex, "Failed to kill process {Pid}", kvp.Key);
            }
    }

    /// <summary>
    ///     Initializes a Windows Job Object with kill-on-close behavior.
    /// </summary>
    private void InitializeWindowsJobObject()
    {
        try
        {
            _jobHandle = CreateJobObject(IntPtr.Zero, null);
            if (_jobHandle.IsInvalid)
            {
                _logger?.LogDebug("Failed to create job object");
                return;
            }

            var info = new JOBOBJECT_EXTENDED_LIMIT_INFORMATION
            {
                BasicLimitInformation = new JOBOBJECT_BASIC_LIMIT_INFORMATION
                {
                    LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
                }
            };

            var length = Marshal.SizeOf(typeof(JOBOBJECT_EXTENDED_LIMIT_INFORMATION));
            var infoPtr = Marshal.AllocHGlobal(length);

            try
            {
                Marshal.StructureToPtr(info, infoPtr, false);

                if (!SetInformationJobObject(
                        _jobHandle,
                        JobObjectInfoType.ExtendedLimitInformation,
                        infoPtr,
                        (uint)length))
                {
                    var error = Marshal.GetLastWin32Error();
                    _logger?.LogDebug("Failed to set job object information, error: {Error}", error);
                    _jobHandle.Dispose();
                    _jobHandle = null;
                    return;
                }

                _jobObjectValid = true;
                _logger?.LogDebug("Windows Job Object initialized for process cleanup");
            }
            finally
            {
                Marshal.FreeHGlobal(infoPtr);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Failed to initialize Windows Job Object");
            _jobHandle?.Dispose();
            _jobHandle = null;
        }
    }

    #region Windows P/Invoke

    /// <summary>
    ///     Windows Job Object limit flag that causes all processes in the job to be
    ///     terminated when the last handle to the job object is closed.
    ///     This ensures child processes are cleaned up even if the parent crashes.
    /// </summary>
    /// <remarks>
    ///     See: https://docs.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-jobobject_basic_limit_information
    /// </remarks>
    private const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000;

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateJobObject(IntPtr lpJobAttributes, string? lpName);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool SetInformationJobObject(
        SafeHandle hJob,
        JobObjectInfoType infoType,
        IntPtr lpJobObjectInfo,
        uint cbJobObjectInfoLength);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool AssignProcessToJobObject(SafeHandle hJob, SafeHandle hProcess);

    private enum JobObjectInfoType
    {
        ExtendedLimitInformation = 9
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JOBOBJECT_BASIC_LIMIT_INFORMATION
    {
        public long PerProcessUserTimeLimit;
        public long PerJobUserTimeLimit;
        public uint LimitFlags;
        public UIntPtr MinimumWorkingSetSize;
        public UIntPtr MaximumWorkingSetSize;
        public uint ActiveProcessLimit;
        public UIntPtr Affinity;
        public uint PriorityClass;
        public uint SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IO_COUNTERS
    {
        public ulong ReadOperationCount;
        public ulong WriteOperationCount;
        public ulong OtherOperationCount;
        public ulong ReadTransferCount;
        public ulong WriteTransferCount;
        public ulong OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    {
        public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
        public IO_COUNTERS IoInfo;
        public UIntPtr ProcessMemoryLimit;
        public UIntPtr JobMemoryLimit;
        public UIntPtr PeakProcessMemoryUsed;
        public UIntPtr PeakJobMemoryUsed;
    }

    #endregion
}
