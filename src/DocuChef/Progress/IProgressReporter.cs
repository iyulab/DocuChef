namespace DocuChef.Progress;

/// <summary>
/// Interface for reporting progress during document processing
/// </summary>
public interface IProgressReporter
{
    /// <summary>
    /// Reports progress with a percentage (0-100)
    /// </summary>
    /// <param name="percentage">Progress percentage (0-100)</param>
    /// <param name="message">Progress message</param>
    void ReportProgress(int percentage, string message);
}

/// <summary>
/// Progress phases for PowerPoint processing
/// </summary>
public enum ProcessingPhase
{
    /// <summary>
    /// Template analysis phase
    /// </summary>
    TemplateAnalysis,
    
    /// <summary>
    /// Alias transformation phase
    /// </summary>
    AliasTransformation,
    
    /// <summary>
    /// Plan generation phase
    /// </summary>
    PlanGeneration,
    
    /// <summary>
    /// Expression binding phase
    /// </summary>
    ExpressionBinding,
    
    /// <summary>
    /// Data binding phase
    /// </summary>
    DataBinding,
    
    /// <summary>
    /// Function processing phase
    /// </summary>
    FunctionProcessing,
    
    /// <summary>
    /// Finalization phase
    /// </summary>
    Finalization
}

/// <summary>
/// Progress information for PowerPoint processing
/// </summary>
public class ProcessingProgress
{
    /// <summary>
    /// Current processing phase
    /// </summary>
    public ProcessingPhase Phase { get; set; }
    
    /// <summary>
    /// Overall progress percentage (0-100)
    /// </summary>
    public int OverallPercentage { get; set; }
    
    /// <summary>
    /// Current phase progress percentage (0-100)
    /// </summary>
    public int PhasePercentage { get; set; }
    
    /// <summary>
    /// Current step within the phase
    /// </summary>
    public int CurrentStep { get; set; }
    
    /// <summary>
    /// Total steps in the current phase
    /// </summary>
    public int TotalSteps { get; set; }
    
    /// <summary>
    /// Progress message
    /// </summary>
    public string Message { get; set; } = string.Empty;
    
    /// <summary>
    /// Additional details
    /// </summary>
    public string Details { get; set; } = string.Empty;
}

/// <summary>
/// Action delegate for progress reporting
/// </summary>
/// <param name="progress">Progress information</param>
public delegate void ProgressCallback(ProcessingProgress progress);