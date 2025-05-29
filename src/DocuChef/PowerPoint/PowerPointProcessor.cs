namespace DocuChef.PowerPoint;

/// <summary>
/// Simple data context for PowerPoint processing
/// </summary>
internal class PowerPointContext
{
    public Dictionary<string, object> Variables { get; set; } = new();
    public Dictionary<string, Func<object>> GlobalVariables { get; set; } = new();
    public Dictionary<string, object> Functions { get; set; } = new();
}

/// <summary>
/// Processes PowerPoint templates with variable optimization
/// </summary>
internal class PowerPointProcessor
{
    private readonly PowerPointContext _context;

    public PowerPointProcessor(PowerPointContext context)
    {
        _context = context ?? throw new ArgumentNullException(nameof(context));
    }

    /// <summary>
    /// Prepare variables dictionary
    /// </summary>
    internal Dictionary<string, object> PrepareVariables()
    {
        var variables = new Dictionary<string, object>(_context.Variables);

        // Add global variables
        foreach (var globalVar in _context.GlobalVariables)
        {
            variables[globalVar.Key] = globalVar.Value();
        }

        // Add PowerPoint functions
        foreach (var function in _context.Functions)
        {
            variables[$"ppt.{function.Key}"] = function.Value;
        }

        return variables;
    }

    /// <summary>
    /// Prepare variables dictionary with filtering based on slide expressions (optimized)
    /// </summary>
    internal Dictionary<string, object> PrepareVariablesForSlide(SlidePart slidePart)
    {
        // Extract all expressions used in this slide
        var usedExpressions = ExtractExpressionsFromSlide(slidePart);
        
        if (!usedExpressions.Any())
        {
            Logger.Debug("No expressions found in slide, returning minimal variables");
            return new Dictionary<string, object> { ["_context"] = _context };
        }

        // Extract variable names from expressions
        var requiredVariables = ExtractVariableNamesFromExpressions(usedExpressions);
        
        Logger.Debug($"Slide requires {requiredVariables.Count} variables: {string.Join(", ", requiredVariables)}");

        var variables = new Dictionary<string, object>();

        // Add context first
        variables["_context"] = _context;

        // Add only required base variables
        foreach (var variableName in requiredVariables)
        {
            if (_context.Variables.TryGetValue(variableName, out var value))
            {
                variables[variableName] = value;
                Logger.Debug($"Added variable: {variableName}");
            }
        }

        // Add required global variables
        foreach (var globalVar in _context.GlobalVariables)
        {
            if (requiredVariables.Contains(globalVar.Key))
            {
                variables[globalVar.Key] = globalVar.Value();
                Logger.Debug($"Added global variable: {globalVar.Key}");
            }
        }

        // Add required PowerPoint functions
        foreach (var function in _context.Functions)
        {
            string functionKey = $"ppt.{function.Key}";
            if (usedExpressions.Any(expr => expr.Contains($"ppt.{function.Key}")))
            {
                variables[functionKey] = function.Value;
                Logger.Debug($"Added function: {functionKey}");
            }
        }

        Logger.Debug($"Prepared {variables.Count} variables for slide (filtered from {_context.Variables.Count + _context.GlobalVariables.Count + _context.Functions.Count} total)");
        return variables;
    }

    /// <summary>
    /// Extract all expressions from a slide
    /// </summary>
    private HashSet<string> ExtractExpressionsFromSlide(SlidePart slidePart)
    {
        var expressions = new HashSet<string>();

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (shape.TextBody == null)
                continue;

            foreach (var text in shape.TextBody.Descendants<A.Text>())
            {
                if (string.IsNullOrEmpty(text.Text))
                    continue;                // Extract expressions using regex directly
                var regex = new Regex(@"\$\{([^}]+)\}", RegexOptions.Compiled);
                var matches = regex.Matches(text.Text);
                foreach (Match match in matches)
                {
                    expressions.Add(match.Groups[1].Value);
                }
            }
        }

        return expressions;
    }

    /// <summary>
    /// Extract variable names from expressions
    /// </summary>
    private HashSet<string> ExtractVariableNamesFromExpressions(HashSet<string> expressions)
    {
        var variableNames = new HashSet<string>();

        foreach (var expression in expressions)
        {
            // Remove ${} wrapper
            var cleanExpr = expression.TrimStart('{', '$').TrimEnd('}');
            
            // Handle different expression types
            if (cleanExpr.StartsWith("ppt."))
            {
                // PowerPoint functions - extract variable references from parameters
                var paramMatch = Regex.Match(cleanExpr, @"ppt\.\w+\(([^)]*)\)");
                if (paramMatch.Success)
                {
                    var parameters = paramMatch.Groups[1].Value;
                    var paramVariables = ExtractVariableNamesFromParameters(parameters);
                    foreach (var paramVar in paramVariables)
                    {
                        variableNames.Add(paramVar);
                    }
                }
            }
            else if (cleanExpr.Contains('[') && cleanExpr.Contains(']'))
            {
                // Array indexing: Items[0].Name, item[0], etc.
                var arrayMatch = Regex.Match(cleanExpr, @"(\w+)\[");
                if (arrayMatch.Success)
                {
                    string arrayName = arrayMatch.Groups[1].Value;
                    // Handle case-insensitive "item" -> "Items" mapping
                    if (string.Equals(arrayName, "item", StringComparison.OrdinalIgnoreCase))
                    {
                        variableNames.Add("Items");
                    }
                    else
                    {
                        variableNames.Add(arrayName);
                    }
                }
            }
            else if (cleanExpr.Contains('.'))
            {
                // Property access: object.property
                var parts = cleanExpr.Split('.');
                variableNames.Add(parts[0]);
            }
            else if (cleanExpr.Contains(':'))
            {
                // Date formatting: Date:yyyy-MM-dd
                var parts = cleanExpr.Split(':');
                variableNames.Add(parts[0]);
            }
            else
            {
                // Simple variable: Title, Subtitle, etc.
                variableNames.Add(cleanExpr);
            }
        }

        return variableNames;
    }

    /// <summary>
    /// Extract variable names from function parameters
    /// </summary>
    private HashSet<string> ExtractVariableNamesFromParameters(string parameters)
    {
        var variableNames = new HashSet<string>();
        
        if (string.IsNullOrEmpty(parameters))
            return variableNames;

        // Simple parameter parsing - extract quoted and unquoted identifiers
        var paramParts = parameters.Split(',');
        foreach (var part in paramParts)
        {
            var cleanPart = part.Trim();
            
            // Skip literals
            if (cleanPart.StartsWith("\"") && cleanPart.EndsWith("\""))
                continue;
            if (cleanPart.StartsWith("'") && cleanPart.EndsWith("'"))
                continue;
            if (int.TryParse(cleanPart, out _) || bool.TryParse(cleanPart, out _))
                continue;

            // Extract variable name (before : if parameter assignment)
            if (cleanPart.Contains(':'))
            {
                var valuePart = cleanPart.Split(':')[1].Trim();
                if (!valuePart.StartsWith("\"") && !valuePart.StartsWith("'") && !int.TryParse(valuePart, out _))
                {
                    variableNames.Add(valuePart);
                }
            }
            else
            {
                // Direct variable reference
                variableNames.Add(cleanPart);
            }
        }

        return variableNames;
    }
}