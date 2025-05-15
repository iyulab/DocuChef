using System.Text.RegularExpressions;
using DollarSignEngine;
using System.Globalization;

namespace DocuChef.PowerPoint.DollarSignEngine;

/// <summary>
/// DollarSignEngine adapter for processing expressions in PowerPoint templates
/// </summary>
internal class ExpressionEvaluator
{
    private readonly DollarSignOptions _options;
    private readonly CultureInfo _cultureInfo;

    /// <summary>
    /// Initializes a new instance of ExpressionEvaluator
    /// </summary>
    public ExpressionEvaluator(CultureInfo cultureInfo = null)
    {
        _cultureInfo = cultureInfo ?? CultureInfo.CurrentCulture;

        _options = new DollarSignOptions
        {
            CultureInfo = _cultureInfo,
            SupportDollarSignSyntax = true,  // Use ${variable} syntax as per PPT guidelines
            ThrowOnMissingParameter = false, // Don't throw on missing params, show placeholder instead
            EnableDebugLogging = false,      // Disable debug logging by default
            VariableResolver = CustomVariableResolver // Custom resolver for PowerPoint functions and array indices
        };
    }

    /// <summary>
    /// Evaluates an expression synchronously
    /// </summary>
    public object Evaluate(string expression, Dictionary<string, object> variables)
    {
        try
        {
            // If already wrapped in ${...}, evaluate directly
            if (expression.StartsWith("${") && expression.EndsWith("}"))
            {
                var result = DollarSign.EvalAsync(expression, variables, _options).GetAwaiter().GetResult();
                return result;
            }

            // Otherwise, wrap it for evaluation
            string wrappedExpr = "${" + expression + "}";
            var evalResult = DollarSign.EvalAsync(wrappedExpr, variables, _options).GetAwaiter().GetResult();
            return evalResult;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error evaluating expression '{expression}': {ex.Message}", ex);
            return $"[Error: {ex.Message}]";
        }
    }

    /// <summary>
    /// Custom variable resolver for PowerPoint functions and array indices
    /// </summary>
    private object CustomVariableResolver(string expression, object parameters)
    {
        // Convert parameters to dictionary if not already
        Dictionary<string, object> variables;
        if (parameters is Dictionary<string, object> dict)
        {
            variables = dict;
        }
        else
        {
            variables = new Dictionary<string, object>();

            // If parameters is an object, extract its properties
            if (parameters != null)
            {
                var props = parameters.GetType().GetProperties();
                foreach (var prop in props)
                {
                    if (prop.CanRead)
                    {
                        try
                        {
                            var value = prop.GetValue(parameters);
                            variables[prop.Name] = value;
                        }
                        catch
                        {
                            // Skip properties that throw exceptions
                        }
                    }
                }
            }
        }

        // Handle PowerPoint functions (ppt.)
        if (expression.StartsWith("ppt."))
        {
            return HandlePowerPointFunction(expression, variables);
        }

        // Handle array indexing expressions: item[0].Property and Items[0].Property
        var arrayIndexMatch = Regex.Match(expression, @"^([\w]+)\[(\d+)\](\.(\w+))?$");
        if (arrayIndexMatch.Success)
        {
            string arrayName = arrayIndexMatch.Groups[1].Value; // e.g., item or Items
            int index = int.Parse(arrayIndexMatch.Groups[2].Value); // e.g., 0, 1, 2
            string propPath = arrayIndexMatch.Groups[4].Success ? arrayIndexMatch.Groups[4].Value : null; // e.g., Id, Name, null

            Logger.Debug($"[EVAL] Array expression: arrayName={arrayName}, index={index}, propPath={propPath}");

            // First check if this is a case-insensitive match for "Items"
            string normalizedName = null;
            if (string.Equals(arrayName, "item", StringComparison.OrdinalIgnoreCase))
            {
                normalizedName = "Items";
            }

            // Try both original and normalized names
            object arrayObj = null;
            if (normalizedName != null && variables.TryGetValue(normalizedName, out var normalizedObj))
            {
                arrayObj = normalizedObj;
                arrayName = normalizedName;
            }
            else if (variables.TryGetValue(arrayName, out var originalObj))
            {
                arrayObj = originalObj;
            }

            if (arrayObj != null)
            {
                Logger.Debug($"[EVAL] Found array '{arrayName}' in variables");

                // Handle different collection types
                if (arrayObj is IList list)
                {
                    Logger.Debug($"[EVAL] Array is IList with count={list.Count}");
                    if (index >= 0 && index < list.Count)
                    {
                        var item = list[index];
                        Logger.Debug($"[EVAL] Retrieved item at index {index}: {item?.GetType().Name}");

                        // Return the item or its property
                        if (string.IsNullOrEmpty(propPath))
                            return item;

                        // Get property value
                        var property = item?.GetType().GetProperty(propPath);
                        if (property != null)
                        {
                            var propValue = property.GetValue(item);
                            Logger.Debug($"[EVAL] Property '{propPath}' value: {propValue}");
                            return propValue;
                        }
                        else
                        {
                            Logger.Warning($"[EVAL] Property '{propPath}' not found on item type {item?.GetType().Name}");
                        }
                    }
                    else
                    {
                        Logger.Warning($"[EVAL] Index {index} out of range for array with count {list.Count}");
                    }
                }
                else if (arrayObj is Array array)
                {
                    Logger.Debug($"[EVAL] Array is Array with length={array.Length}");
                    if (index >= 0 && index < array.Length)
                    {
                        var item = array.GetValue(index);
                        Logger.Debug($"[EVAL] Retrieved item at index {index}: {item?.GetType().Name}");

                        // Return the item or its property
                        if (string.IsNullOrEmpty(propPath))
                            return item;

                        // Get property value
                        var property = item?.GetType().GetProperty(propPath);
                        if (property != null)
                        {
                            var propValue = property.GetValue(item);
                            Logger.Debug($"[EVAL] Property '{propPath}' value: {propValue}");
                            return propValue;
                        }
                    }
                }
                else if (arrayObj is IEnumerable enumerable)
                {
                    Logger.Debug($"[EVAL] Array is IEnumerable");
                    // Try to get the item by index
                    int currentIndex = 0;
                    foreach (var item in enumerable)
                    {
                        if (currentIndex == index)
                        {
                            Logger.Debug($"[EVAL] Retrieved item at index {index}: {item?.GetType().Name}");

                            // Return the item or its property
                            if (string.IsNullOrEmpty(propPath))
                                return item;

                            // Get property value
                            var property = item?.GetType().GetProperty(propPath);
                            if (property != null)
                            {
                                var propValue = property.GetValue(item);
                                Logger.Debug($"[EVAL] Property '{propPath}' value: {propValue}");
                                return propValue;
                            }
                        }
                        currentIndex++;
                    }
                }
            }
            else
            {
                Logger.Warning($"[EVAL] Array '{arrayName}' not found in variables");
            }
        }

        // For other expressions, let DollarSignEngine handle them
        return null;
    }

    /// <summary>
    /// Handles PowerPoint specific functions (ppt.Image, ppt.Chart, ppt.Table)
    /// </summary>
    private object HandlePowerPointFunction(string expression, Dictionary<string, object> variables)
    {
        // Parse function expression: ppt.Function("arg", param1: value1, param2: value2)
        var match = Regex.Match(expression, @"ppt\.(\w+)\((.+)??\)");
        if (!match.Success)
        {
            Logger.Warning($"Invalid PowerPoint function format: {expression}");
            return $"[Invalid function: {expression}]";
        }

        string functionName = match.Groups[1].Value;
        string argsString = match.Groups[2].Success ? match.Groups[2].Value : "";

        // Parse arguments
        string[] args = ParseFunctionArguments(argsString);

        // Look up the PowerPoint function
        if (variables.TryGetValue($"ppt.{functionName}", out var funcObj) &&
            funcObj is PowerPointFunction function)
        {
            // Look up PowerPoint context
            PowerPointContext context = null;
            if (variables.TryGetValue("_context", out var ctxObj) &&
                ctxObj is PowerPointContext ctx)
            {
                context = ctx;
            }
            else
            {
                // If no context (rare case), create a temporary one
                context = new PowerPointContext
                {
                    Variables = new Dictionary<string, object>(variables)
                };
            }

            try
            {
                // Execute the function
                return function.Execute(context, null, args);
            }
            catch (Exception ex)
            {
                Logger.Error($"Error executing function {functionName}: {ex.Message}", ex);
                return $"[Error in {functionName}: {ex.Message}]";
            }
        }

        return $"[Unknown function: ppt.{functionName}]";
    }

    /// <summary>
    /// Simple parsing of function arguments
    /// </summary>
    private string[] ParseFunctionArguments(string argsString)
    {
        if (string.IsNullOrEmpty(argsString))
            return Array.Empty<string>();

        var args = new List<string>();
        bool inQuotes = false;
        int startIndex = 0;

        for (int i = 0; i < argsString.Length; i++)
        {
            char c = argsString[i];

            if (c == '"' && (i == 0 || argsString[i - 1] != '\\'))
            {
                inQuotes = !inQuotes;
            }
            else if (c == ',' && !inQuotes)
            {
                args.Add(argsString.Substring(startIndex, i - startIndex).Trim());
                startIndex = i + 1;
            }
        }

        // Add the last argument
        if (startIndex < argsString.Length)
        {
            args.Add(argsString.Substring(startIndex).Trim());
        }

        // Clean up arguments
        for (int i = 0; i < args.Count; i++)
        {
            var arg = args[i].Trim();

            // Remove quotes from string arguments
            if (arg.StartsWith("\"") && arg.EndsWith("\"") && arg.Length > 1)
            {
                arg = arg.Substring(1, arg.Length - 2)
                    .Replace("\\\"", "\"")
                    .Replace("\\\\", "\\");
            }

            args[i] = arg;
        }

        return args.ToArray();
    }
}