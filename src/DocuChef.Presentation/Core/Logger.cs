namespace DocuChef.Presentation.Core;

/// <summary>
/// Provides logging functionality for the presentation library
/// </summary>
public static class Logger
{
    /// <summary>
    /// Log level for controlling output verbosity
    /// </summary>
    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }

    /// <summary>
    /// Current log level
    /// </summary>
    public static LogLevel CurrentLevel { get; set; } = LogLevel.Info;

    /// <summary>
    /// Logs a debug message when appropriate log level is set
    /// </summary>
    public static void Debug(string message)
    {
        if (CurrentLevel <= LogLevel.Debug)
            LogMessage("DEBUG", message);
    }

    /// <summary>
    /// Logs an info message
    /// </summary>
    public static void Info(string message)
    {
        if (CurrentLevel <= LogLevel.Info)
            LogMessage("INFO", message);
    }

    /// <summary>
    /// Logs a warning message
    /// </summary>
    public static void Warning(string message)
    {
        if (CurrentLevel <= LogLevel.Warning)
            LogMessage("WARNING", message);
    }

    /// <summary>
    /// Logs an error message
    /// </summary>
    public static void Error(string message, Exception ex = null)
    {
        if (CurrentLevel <= LogLevel.Error)
        {
            LogMessage("ERROR", message);

            if (ex != null)
            {
                LogMessage("ERROR", $"Exception: {ex.Message}");
                LogMessage("ERROR", $"Stack trace: {ex.StackTrace}");
            }
        }
    }

    /// <summary>
    /// Logs a message with a consistent format
    /// </summary>
    private static void LogMessage(string level, string message)
    {
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        string formattedMessage = $"[DocuChef {level}] {timestamp} - {message}";

        // Output to debug window
        System.Diagnostics.Debug.WriteLine(formattedMessage);

        // Output to console
        Console.WriteLine(formattedMessage);
    }
}
