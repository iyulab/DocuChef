namespace DocuChef.PowerPoint;

/// <summary>
/// Represents a custom function for PowerPoint processing
/// </summary>
public class PowerPointFunction
{
    /// <summary>
    /// Function name
    /// </summary>
    public required string Name { get; set; }

    /// <summary>
    /// Function handler
    /// </summary>
    public required Func<PowerPointContext, object?, string[]?, object?> Handler { get; set; }

    /// <summary>
    /// Function description
    /// </summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>
    /// Creates a new PowerPoint function
    /// </summary>
    public PowerPointFunction() { }

    /// <summary>
    /// Execute the function
    /// </summary>
    public object? Execute(PowerPointContext context, object? value, string[]? parameters)
    {
        if (Handler == null)
        {
            Logger.Warning($"No handler defined for PowerPoint function '{Name}'");
            return $"[Error: Function '{Name}' has no implementation]";
        }

        try
        {
            Logger.Debug($"Executing PowerPoint function '{Name}' with {parameters?.Length ?? 0} parameters");
            var result = Handler(context, value, parameters);
            Logger.Debug($"Function '{Name}' executed successfully");
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error executing PowerPoint function '{Name}'", ex);
            return $"[Error in function '{Name}': {ex.Message}]";
        }
    }
}