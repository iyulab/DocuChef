using DocuChef.Logging;

namespace DocuChef.Progress;

/// <summary>
/// Default implementation of progress reporter
/// </summary>
public class ProgressReporter : IProgressReporter
{
    private readonly ProgressCallback? _callback;
    private readonly bool _enableLogging;
    private ProcessingProgress _currentProgress;

    /// <summary>
    /// Creates a new progress reporter with optional callback
    /// </summary>
    /// <param name="callback">Optional callback for progress updates</param>
    /// <param name="enableLogging">Whether to enable logging</param>
    public ProgressReporter(ProgressCallback? callback = null, bool enableLogging = true)
    {
        _callback = callback;
        _enableLogging = enableLogging;
        _currentProgress = new ProcessingProgress();
    }

    /// <summary>
    /// Reports progress with a percentage and message
    /// </summary>
    /// <param name="percentage">Progress percentage (0-100)</param>
    /// <param name="message">Progress message</param>
    public void ReportProgress(int percentage, string message)
    {
        _currentProgress.OverallPercentage = Math.Max(0, Math.Min(100, percentage));
        _currentProgress.Message = message ?? string.Empty;

        if (_enableLogging)
        {
            Logger.Info($"Progress: {percentage}% - {message}");
        }

        _callback?.Invoke(_currentProgress);
    }

    /// <summary>
    /// Reports progress for a specific phase
    /// </summary>
    /// <param name="phase">Current processing phase</param>
    /// <param name="phasePercentage">Phase progress percentage (0-100)</param>
    /// <param name="overallPercentage">Overall progress percentage (0-100)</param>
    /// <param name="message">Progress message</param>
    /// <param name="details">Additional details</param>
    public void ReportPhaseProgress(ProcessingPhase phase, int phasePercentage, int overallPercentage, string message, string details = "")
    {
        _currentProgress.Phase = phase;
        _currentProgress.PhasePercentage = Math.Max(0, Math.Min(100, phasePercentage));
        _currentProgress.OverallPercentage = Math.Max(0, Math.Min(100, overallPercentage));
        _currentProgress.Message = message ?? string.Empty;
        _currentProgress.Details = details ?? string.Empty;

        if (_enableLogging)
        {
            Logger.Info($"Phase {phase}: {phasePercentage}% (Overall: {overallPercentage}%) - {message}");
            if (!string.IsNullOrEmpty(details))
            {
                Logger.Debug($"Details: {details}");
            }
        }

        _callback?.Invoke(_currentProgress);
    }

    /// <summary>
    /// Reports progress for a specific step within a phase
    /// </summary>
    /// <param name="phase">Current processing phase</param>
    /// <param name="currentStep">Current step number</param>
    /// <param name="totalSteps">Total number of steps</param>
    /// <param name="overallPercentage">Overall progress percentage (0-100)</param>
    /// <param name="message">Progress message</param>
    /// <param name="details">Additional details</param>
    public void ReportStepProgress(ProcessingPhase phase, int currentStep, int totalSteps, int overallPercentage, string message, string details = "")
    {
        var phasePercentage = totalSteps > 0 ? (currentStep * 100 / totalSteps) : 0;
        
        _currentProgress.Phase = phase;
        _currentProgress.CurrentStep = currentStep;
        _currentProgress.TotalSteps = totalSteps;
        _currentProgress.PhasePercentage = phasePercentage;
        _currentProgress.OverallPercentage = Math.Max(0, Math.Min(100, overallPercentage));
        _currentProgress.Message = message ?? string.Empty;
        _currentProgress.Details = details ?? string.Empty;

        if (_enableLogging)
        {
            Logger.Info($"Phase {phase}: Step {currentStep}/{totalSteps} (Overall: {overallPercentage}%) - {message}");
            if (!string.IsNullOrEmpty(details))
            {
                Logger.Debug($"Details: {details}");
            }
        }

        _callback?.Invoke(_currentProgress);
    }

    /// <summary>
    /// Gets the current progress state
    /// </summary>
    public ProcessingProgress GetCurrentProgress()
    {
        return _currentProgress;
    }
}