using System.Text.RegularExpressions;
using DollarSignEngine;
using DocuChef.Logging;
using DocuChef.Presentation.Models;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// Handles data binding by converting template expressions to DollarSign format and evaluating them.
/// Supports context operators (>) for nested property access.
/// </summary>
public class DataBinder
{
    private static readonly Regex ContextOperatorRegex = new(@"(\w+)>([^}\s]+)", RegexOptions.Compiled); private static readonly Regex DollarSignExpressionRegex = new(@"\$\{([^}]+)\}", RegexOptions.Compiled);    /// <summary>
                                                                                                                                                                                                                    /// Binds data to template expressions in the given text.
                                                                                                                                                                                                                    /// Converts context operators (>) to underscore notation (___) and evaluates using DollarSignEngine.
                                                                                                                                                                                                                    /// </summary>
                                                                                                                                                                                                                    /// <param name="template">Template text containing expressions</param>
                                                                                                                                                                                                                    /// <param name="data">Data object to bind</param>
                                                                                                                                                                                                                    /// <param name="indexOffset">Index offset for array expressions</param>
                                                                                                                                                                                                                    /// <returns>Text with all expressions evaluated</returns>
    public string BindData(string template, object data, int indexOffset = 0)
    {
        if (string.IsNullOrEmpty(template) || data == null)
            return template ?? string.Empty;
        try
        {
            // Pre-process nested expressions first
            var preprocessedTemplate = PreprocessNestedExpressions(template, data, indexOffset);
            Logger.Debug($"DataBinder: Original template: '{template}'");
            Logger.Debug($"DataBinder: Preprocessed template: '{preprocessedTemplate}'");

            // Convert template to DollarSign format
            var dollarSignTemplate = ConvertToTemplate(preprocessedTemplate);
            Logger.Debug($"DataBinder: Converted template to '{dollarSignTemplate}'");            // Create context variables for the data
            var variables = CreateContextOperatorVariables(data);

            // Add PPT functions to variables - reuse existing instance if available
            DocuChef.Presentation.Functions.PPTFunctions pptFunctions;
            if (data is Dictionary<string, object> dataDict &&
                dataDict.TryGetValue("ppt", out var existingPpt) &&
                existingPpt is DocuChef.Presentation.Functions.PPTFunctions existingPptFunctions)
            {
                // Reuse existing PPTFunctions instance (preserves image cache)
                pptFunctions = existingPptFunctions;
                Logger.Debug($"DataBinder: Reusing existing PPTFunctions instance with {existingPptFunctions.GetAllImageCache().Count} cached images");
            }
            else
            {
                // Create new PPTFunctions instance
                pptFunctions = new DocuChef.Presentation.Functions.PPTFunctions(variables);
                Logger.Debug($"DataBinder: Created new PPTFunctions instance");
            }
            variables["ppt"] = pptFunctions;

            Logger.Debug($"DataBinder: Created {variables.Count} variables:");
            foreach (var kvp in variables)
            {
                if (kvp.Value is System.Collections.IEnumerable enumerable && !(kvp.Value is string))
                {
                    var list = enumerable.Cast<object>().ToList();
                    Logger.Debug($"  {kvp.Key} = [{string.Join(", ", list.Take(3).Select(x => x?.ToString() ?? "null"))}] (count: {list.Count})");
                }
                else
                {
                    Logger.Debug($"  {kvp.Key} = {kvp.Value}");
                }
            }// Apply index offset to array expressions
            if (indexOffset > 0)
            {
                dollarSignTemplate = ApplyIndexOffset(dollarSignTemplate, indexOffset);
            }

            // Configure DollarSign options
            var options = new DollarSignOptions
            {
                SupportDollarSignSyntax = true,
                ThrowOnError = false,
                UseCache = true,
                CultureInfo = System.Globalization.CultureInfo.CurrentCulture
            };

            // Evaluate the template with variables - using EvalAsync synchronously
            var result = DollarSign.EvalAsync(dollarSignTemplate, variables, options).GetAwaiter().GetResult();
            // DEBUG: Test simple template evaluation
            if (dollarSignTemplate.Contains("Title"))
            {
                Logger.Debug($"DataBinder: DEBUG - Testing simple Title evaluation");
                var testResult = DollarSign.EvalAsync("${Title}", variables, options).GetAwaiter().GetResult();
                Logger.Debug($"DataBinder: DEBUG - Simple '${{Title}}' evaluated to '{testResult}'");

                // Test if the variable exists in the dictionary
                if (variables.ContainsKey("Title"))
                {
                    Logger.Debug($"DataBinder: DEBUG - Title variable exists: '{variables["Title"]}'");
                }
                else
                {
                    Logger.Debug($"DataBinder: DEBUG - Title variable missing! Available keys: {string.Join(", ", variables.Keys.Take(10))}");
                }
            }

            Logger.Debug($"DataBinder: Template '{dollarSignTemplate}' evaluated to '{result}'");
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"DataBinder: Error binding data to template '{template}': {ex.Message}");
            return template; // Return original template on error
        }
    }

    /// <summary>
    /// Converts template expressions to DollarSign format.
    /// Transforms context operators (>) to underscore notation (___).
    /// Example: Category>Products[0].Name → ${Category___Products[0].Name}
    /// </summary>
    private string ConvertToTemplate(string template)
    {
        if (string.IsNullOrEmpty(template))
            return string.Empty;

        var result = template;

        // Handle context operators - convert the entire expression
        var contextMatches = ContextOperatorRegex.Matches(result);

        foreach (Match match in contextMatches)
        {
            var fullExpression = match.Value; // e.g., "Categories>Items[0].Name"
            var converted = fullExpression.Replace(">", "___");

            // Wrap in DollarSign syntax if not already wrapped
            if (!result.Contains($"${{{converted}}}"))
            {
                result = result.Replace(fullExpression, $"${{{converted}}}");
            }
        }        // Convert array indexing with property access: Items[0].Id -> Items_0___Id
        // Also handle formatting specifiers: Items[0].Price:C0 -> Items_0___Price:C0
        var arrayPropertyRegex = new Regex(@"\$\{([a-zA-Z_]\w*)\[(\d+)\]\.([a-zA-Z_]\w*)(:[\w\-\+#,\.]*?)?\}", RegexOptions.Compiled);
        result = arrayPropertyRegex.Replace(result, @"${$1_$2___$3$4}");

        // Skip bare property wrapping if the template already contains properly formatted expressions
        // This prevents interfering with function calls like ppt.Image(LogoPath)
        if (result.Contains("${") && result.Contains("}"))
        {
            Logger.Debug($"DataBinder: Template already contains ${{}} expressions, skipping bare property wrapping");
            return result;
        }

        // Handle other expressions that might not have context operators
        // Look for bare property names that should be wrapped
        var barePropertyRegex = new Regex(@"\b([A-Z]\w*(?:\[\d+\])?(?:\.\w+)*)\b", RegexOptions.Compiled);
        var bareMatches = barePropertyRegex.Matches(result);

        foreach (Match match in bareMatches)
        {
            var property = match.Value;

            // Skip if already processed, is a DollarSign expression, or contains underscore notation
            if (result.Contains($"${{{property}}}") ||
                property.Contains("___") ||
                result.Substring(Math.Max(0, match.Index - 2), Math.Min(result.Length - Math.Max(0, match.Index - 2), property.Length + 4)).Contains($"${{{property}}}"))
                continue;

            // Skip common words that shouldn't be treated as properties
            if (IsCommonWord(property))
                continue;

            // Wrap in DollarSign syntax
            result = result.Replace(property, $"${{{property}}}");
        }

        return result;
    }

    /// <summary>
    /// Creates context operator variables for DollarSignEngine.
    /// Flattens nested properties into underscore-separated variables.
    /// </summary>
    private Dictionary<string, object> CreateContextOperatorVariables(object data)
    {
        var variables = new Dictionary<string, object>();

        if (data == null)
            return variables;

        // Add the root data object
        variables["Data"] = data;

        // Create flattened variables for nested properties
        CreateContextVariablesRecursive(data, string.Empty, variables, 0);

        return variables;
    }
    /// <summary>
    /// Recursively creates context variables for nested objects and arrays.
    /// </summary>
    private void CreateContextVariablesRecursive(object obj, string prefix, Dictionary<string, object> variables, int depth)
    {
        if (obj == null || depth > 5) // Prevent infinite recursion
            return;        // Handle Dictionary<string, object> specially
        if (obj is Dictionary<string, object> dictionary)
        {
            foreach (var kvp in dictionary)
            {
                if (kvp.Value == null)
                    continue;

                var propertyName = string.IsNullOrEmpty(prefix)
                    ? kvp.Key
                    : $"{prefix}___{kvp.Key}";

                // Add the property value
                variables[propertyName] = kvp.Value;

                // Handle arrays/collections - need to extract nested properties from array elements
                if (kvp.Value is System.Collections.IEnumerable enumerable && !(kvp.Value is string))
                {
                    CreateContextVariablesForArray(enumerable, propertyName, variables, depth + 1);

                    // For context operators, also extract common properties from array elements
                    var list = enumerable.Cast<object>().ToList();
                    if (list.Count > 0 && list[0] != null)
                    {
                        // Extract common properties from the first element to create context operator paths
                        ExtractArrayElementProperties(list, propertyName, variables, depth + 1);
                    }
                }
                // Handle complex objects
                else if (!IsSimpleType(kvp.Value.GetType()))
                {
                    CreateContextVariablesRecursive(kvp.Value, propertyName, variables, depth + 1);
                }
            }
            return;
        }
        var type = obj.GetType();
        var properties = type.GetProperties();

        foreach (var property in properties)
        {
            try
            {
                var value = property.GetValue(obj);
                if (value == null)
                    continue;

                var propertyName = string.IsNullOrEmpty(prefix)
                    ? property.Name
                    : $"{prefix}___{property.Name}";

                // Add the property value
                variables[propertyName] = value;

                // Handle arrays/collections
                if (value is System.Collections.IEnumerable enumerable && !(value is string))
                {
                    CreateContextVariablesForArray(enumerable, propertyName, variables, depth + 1);

                    // For context operators, also extract common properties from array elements
                    var list = enumerable.Cast<object>().ToList();
                    if (list.Count > 0 && list[0] != null)
                    {
                        // Extract common properties from the first element to create context operator paths
                        ExtractArrayElementProperties(list, propertyName, variables, depth + 1);
                    }
                }
                // Handle complex objects
                else if (!IsSimpleType(value.GetType()))
                {
                    CreateContextVariablesRecursive(value, propertyName, variables, depth + 1);
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"DataBinder: Error processing property '{property.Name}': {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Creates context variables for array/collection elements.
    /// </summary>
    private void CreateContextVariablesForArray(System.Collections.IEnumerable enumerable, string propertyName, Dictionary<string, object> variables, int depth)
    {
        var list = enumerable.Cast<object>().ToList();        // Add indexed access for each element
        for (int i = 0; i < list.Count; i++)
        {
            var element = list[i];
            if (element == null)
                continue;

            var indexedName = $"{propertyName}_{i}";
            variables[indexedName] = element;

            // Recursively process complex objects in arrays
            if (!IsSimpleType(element.GetType()))
            {
                CreateContextVariablesRecursive(element, indexedName, variables, depth + 1);
            }
        }

        // Add the array itself
        variables[propertyName] = list;
    }

    /// <summary>
    /// Applies index offset to array expressions in the template.
    /// </summary>
    public string ApplyIndexOffset(string template, int offset)
    {
        if (offset == 0 || string.IsNullOrEmpty(template))
            return template;

        var indexRegex = new Regex(@"\[(\d+)\]", RegexOptions.Compiled);

        return indexRegex.Replace(template, match =>
        {
            if (int.TryParse(match.Groups[1].Value, out var index))
            {
                var newIndex = index + offset;
                return $"[{newIndex}]";
            }
            return match.Value;
        });
    }

    /// <summary>
    /// Checks if a type is a simple type (primitive, string, DateTime, etc.)
    /// </summary>
    private static bool IsSimpleType(Type type)
    {
        return type.IsPrimitive ||
               type == typeof(string) ||
               type == typeof(DateTime) ||
               type == typeof(DateTimeOffset) ||
               type == typeof(decimal) ||
               type == typeof(Guid) ||
               type == typeof(TimeSpan) ||
               (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>));
    }
    /// <summary>
    /// Checks if a word is a common word that shouldn't be treated as a property.
    /// </summary>
    private static bool IsCommonWord(string word)
    {
        var commonWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Value", "Text", "Title", "Description", "Type", "Size", "Count",
            "Date", "Time", "Year", "Month", "Day", "Hour", "Minute", "Second",
            "True", "False", "Yes", "No", "On", "Off", "Enable", "Disable"
        };

        return commonWords.Contains(word);
    }
    /// <summary>
    /// Resolves an expression and returns the result as a string.
    /// This is a simpler interface for basic expression resolution.
    /// </summary>
    /// <param name="expression">Expression to resolve</param>
    /// <param name="data">Data object to bind</param>
    /// <returns>Resolved expression as string</returns>
    public string ResolveExpression(string expression, object data)
    {
        if (string.IsNullOrEmpty(expression))
            return string.Empty;

        if (data == null)
            return string.Empty;

        // Preprocess PowerPoint function expressions to fix malformed syntax
        expression = PreprocessPowerPointExpressions(expression);

        return BindData(expression, data);
    }

    /// <summary>
    /// Preprocesses PowerPoint expressions to fix common malformed syntax issues
    /// </summary>
    private string PreprocessPowerPointExpressions(string expression)
    {
        if (string.IsNullOrEmpty(expression) || !expression.Contains("ppt."))
            return expression;

        try
        {
            // Fix expressions like ppt.${Image}(${...}) to proper ppt.Image("...")
            expression = System.Text.RegularExpressions.Regex.Replace(
                expression,
                @"ppt\.\$\{(\w+)\}\(\$\{([^}]+)\}\)",
                match =>
                {
                    string functionName = match.Groups[1].Value;
                    string parameter = match.Groups[2].Value;

                    // Clean up complex parameter expressions
                    parameter = CleanParameterExpression(parameter);

                    return $"ppt.{functionName}(\"{parameter}\")";
                },
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            );

            Logger.Debug($"Preprocessed PowerPoint expression: {expression}");
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error preprocessing PowerPoint expression '{expression}': {ex.Message}");
        }

        return expression;
    }

    /// <summary>
    /// Cleans up parameter expressions to extract variable names
    /// </summary>
    private string CleanParameterExpression(string parameter)
    {
        if (string.IsNullOrEmpty(parameter))
            return parameter;

        // Remove casting and complex syntax: ((string)Globals["LogoPath"]) -> LogoPath
        parameter = System.Text.RegularExpressions.Regex.Replace(
            parameter,
            @"\(\(string\)\s*Globals\[\s*[""']([^""']+)[""']\s*\]\)",
            "$1",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase
        );

        // Remove other casting patterns
        parameter = System.Text.RegularExpressions.Regex.Replace(
            parameter,
            @"\(\([^)]+\)\s*([^)]+)\)",
            "$1",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase
        );

        return parameter.Trim();
    }
    /// <summary>
    /// Applies index offset to array expressions in a BindingExpression.
    /// </summary>
    /// <param name="bindingExpression">Binding expression to adjust</param>
    /// <param name="offset">Index offset to apply</param>
    /// <returns>Adjusted BindingExpression</returns>
    public BindingExpression ApplyIndexOffset(BindingExpression bindingExpression, int offset)
    {
        if (bindingExpression == null || offset == 0)
            return bindingExpression ?? new BindingExpression();

        // Create a new BindingExpression with adjusted original expression
        var adjustedExpression = ApplyIndexOffset(bindingExpression.OriginalExpression, offset);
        var adjustedDataPath = ApplyIndexOffset(bindingExpression.DataPath, offset);

        return new BindingExpression
        {
            OriginalExpression = adjustedExpression,
            DataPath = adjustedDataPath,
            FormatSpecifier = bindingExpression.FormatSpecifier,
            UsesContextOperator = bindingExpression.UsesContextOperator,
            IsConditional = bindingExpression.IsConditional,
            IsMethodCall = bindingExpression.IsMethodCall,
            ArrayIndices = new Dictionary<string, int>(bindingExpression.ArrayIndices.ToDictionary(
                kvp => kvp.Key,
                kvp => kvp.Value + offset))
        };
    }

    /// <summary>
    /// Pre-processes nested expressions like ${ppt.Image(${LogoPath})} by evaluating inner expressions first.
    /// This prevents DollarSignEngine from trying to compile invalid C# with nested ${} syntax.
    /// </summary>
    /// <param name="template">Template text with potentially nested expressions</param>
    /// <param name="data">Data object to bind</param>
    /// <param name="indexOffset">Index offset for array expressions</param>
    /// <returns>Template with nested expressions resolved</returns>
    private string PreprocessNestedExpressions(string template, object data, int indexOffset = 0)
    {
        if (string.IsNullOrEmpty(template))
            return string.Empty;

        var result = template;
        var nestedExpressionRegex = new Regex(@"\$\{([^{}]*\$\{[^}]+\}[^{}]*)\}", RegexOptions.Compiled);

        // Keep processing until no more nested expressions are found
        var maxIterations = 10; // Prevent infinite loops
        var iteration = 0;

        while (nestedExpressionRegex.IsMatch(result) && iteration < maxIterations)
        {
            var matches = nestedExpressionRegex.Matches(result);

            foreach (Match match in matches.Cast<Match>().Reverse()) // Process from right to left to avoid index issues
            {
                var fullExpression = match.Value; // e.g., "${ppt.Image(${LogoPath})}"
                var innerContent = match.Groups[1].Value; // e.g., "ppt.Image(${LogoPath})"

                Logger.Debug($"DataBinder: Processing nested expression '{fullExpression}'");

                // First, resolve any inner ${} expressions within this content
                var innerResolved = ResolveInnerExpressions(innerContent, data, indexOffset);

                // Replace the full expression with the resolved version
                var resolvedExpression = $"${{{innerResolved}}}";
                result = result.Substring(0, match.Index) + resolvedExpression + result.Substring(match.Index + match.Length);

                Logger.Debug($"DataBinder: Resolved '{fullExpression}' to '{resolvedExpression}'");
            }

            iteration++;
        }

        return result;
    }

    /// <summary>
    /// Resolves inner ${} expressions within a larger expression.
    /// For example, "ppt.Image(${LogoPath})" becomes "ppt.Image(ActualLogoPathValue)"
    /// </summary>
    private string ResolveInnerExpressions(string content, object data, int indexOffset = 0)
    {
        var innerExpressionRegex = new Regex(@"\$\{([^}]+)\}", RegexOptions.Compiled);
        var matches = innerExpressionRegex.Matches(content);

        var result = content;

        foreach (Match match in matches.Cast<Match>().Reverse()) // Process from right to left
        {
            var expression = match.Groups[1].Value; // e.g., "LogoPath"

            try
            {
                // Create a simple template with just this expression
                var simpleTemplate = $"${{{expression}}}";

                // Convert to DollarSign format and evaluate
                var dollarSignTemplate = ConvertToTemplate(simpleTemplate);
                var variables = CreateContextOperatorVariables(data);

                if (indexOffset > 0)
                {
                    dollarSignTemplate = ApplyIndexOffset(dollarSignTemplate, indexOffset);
                }

                var options = new DollarSignOptions
                {
                    SupportDollarSignSyntax = true,
                    ThrowOnError = false,
                    UseCache = true,
                    CultureInfo = System.Globalization.CultureInfo.CurrentCulture
                };

                var evaluatedValue = DollarSign.EvalAsync(dollarSignTemplate, variables, options).GetAwaiter().GetResult();

                // Replace the ${expression} with the evaluated value
                result = result.Substring(0, match.Index) + evaluatedValue + result.Substring(match.Index + match.Length);

                Logger.Debug($"DataBinder: Resolved inner expression '${expression}' to '{evaluatedValue}'");
            }
            catch (Exception ex)
            {
                Logger.Error($"DataBinder: Error resolving inner expression '${expression}': {ex.Message}");
                // Keep the original expression if evaluation fails
            }
        }

        return result;
    }    /// <summary>
         /// Extracts properties from array elements to create context operator variables.
         /// Creates indexed variables that preserve object type information.
         /// For example: Categories___Items[0], Categories___Items[1], etc.
         /// </summary>
    private void ExtractArrayElementProperties(List<object> arrayElements, string arrayPropertyName, Dictionary<string, object> variables, int depth)
    {
        if (arrayElements == null || arrayElements.Count == 0 || depth > 5)
            return;

        // Get properties from the first element (assuming all elements have similar structure)
        var firstElement = arrayElements[0];
        if (firstElement == null)
            return;

        var elementType = firstElement.GetType();
        var properties = elementType.GetProperties();

        foreach (var property in properties)
        {
            try
            {
                var propertyValue = property.GetValue(firstElement);
                if (propertyValue == null)
                    continue;

                // Create context operator variable for this property across all array elements
                var contextPropertyName = $"{arrayPropertyName}___{property.Name}";

                // Collect this property from all elements in the array
                var propertyValues = new List<object>();
                foreach (var element in arrayElements)
                {
                    if (element != null)
                    {
                        var elementValue = property.GetValue(element);
                        if (elementValue != null)
                        {
                            propertyValues.Add(elementValue);
                        }
                    }
                }

                if (propertyValues.Count > 0)
                {
                    // For arrays, we need to flatten them for context operator access
                    if (propertyValues[0] is System.Collections.IEnumerable enumerable && !(propertyValues[0] is string))
                    {
                        // Flatten all arrays into a single list while preserving object references
                        var flattenedList = new List<object>();
                        foreach (var propValue in propertyValues)
                        {
                            if (propValue is System.Collections.IEnumerable enumPropValue && !(propValue is string))
                            {
                                flattenedList.AddRange(enumPropValue.Cast<object>());
                            }
                        }

                        if (flattenedList.Count > 0)
                        {
                            // Store as List<object> to preserve object references
                            variables[contextPropertyName] = flattenedList;

                            // Also create indexed variables for direct access
                            for (int i = 0; i < flattenedList.Count; i++)
                            {
                                variables[$"{contextPropertyName}[{i}]"] = flattenedList[i];
                            }

                            Logger.Debug($"DataBinder: Created flattened context variable '{contextPropertyName}' with {flattenedList.Count} elements");

                            // Recursively extract properties from the flattened array
                            ExtractArrayElementProperties(flattenedList, contextPropertyName, variables, depth + 1);
                        }
                    }
                    else
                    {
                        // For non-array properties, store as List<object> to preserve object references
                        variables[contextPropertyName] = propertyValues;

                        // Also create indexed variables for direct access
                        for (int i = 0; i < propertyValues.Count; i++)
                        {
                            variables[$"{contextPropertyName}[{i}]"] = propertyValues[i];
                        }

                        Logger.Debug($"DataBinder: Created context variable '{contextPropertyName}' with {propertyValues.Count} elements");

                        // If the property value is a complex object, recurse
                        if (!IsSimpleType(propertyValues[0].GetType()))
                        {
                            CreateContextVariablesRecursive(propertyValues[0], contextPropertyName, variables, depth + 1);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"DataBinder: Error extracting property '{property.Name}' from array elements: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Binds data to template expressions in the given text.
    /// Converts context operators (>) to underscore notation (___) and evaluates using DollarSignEngine.
    /// This overload only creates variables for expressions that are actually used in the template.
    /// </summary>
    /// <param name="template">Template text containing expressions</param>
    /// <param name="data">Data object to bind</param>
    /// <param name="usedExpressions">Set of expressions actually used in the template</param>
    /// <param name="indexOffset">Index offset for array expressions</param>
    /// <returns>Text with all expressions evaluated</returns>
    public string BindData(string template, object data, ISet<string> usedExpressions, int indexOffset = 0)
    {
        if (string.IsNullOrEmpty(template) || data == null)
            return template ?? string.Empty;

        try
        {
            // Pre-process nested expressions first
            var preprocessedTemplate = PreprocessNestedExpressions(template, data, indexOffset);
            Logger.Debug($"DataBinder: Original template: '{template}'");
            Logger.Debug($"DataBinder: Preprocessed template: '{preprocessedTemplate}'");

            // Convert template to DollarSign format
            var dollarSignTemplate = ConvertToTemplate(preprocessedTemplate);
            Logger.Debug($"DataBinder: Converted template to '{dollarSignTemplate}'");            // Create context variables for the data - OPTIMIZED VERSION
            var variables = CreateContextOperatorVariablesFiltered(data, usedExpressions);

            // Add PPT functions to variables - reuse existing instance if available
            DocuChef.Presentation.Functions.PPTFunctions pptFunctions;
            if (data is Dictionary<string, object> dataDict &&
                dataDict.TryGetValue("ppt", out var existingPpt) &&
                existingPpt is DocuChef.Presentation.Functions.PPTFunctions existingPptFunctions)
            {
                // Reuse existing PPTFunctions instance (preserves image cache)
                pptFunctions = existingPptFunctions;
                Logger.Debug($"DataBinder: Reusing existing PPTFunctions instance with {existingPptFunctions.GetAllImageCache().Count} cached images");
            }
            else
            {
                // Create new PPTFunctions instance
                pptFunctions = new DocuChef.Presentation.Functions.PPTFunctions(variables);
                Logger.Debug($"DataBinder: Created new PPTFunctions instance");
            }
            variables["ppt"] = pptFunctions;

            Logger.Debug($"DataBinder: Created {variables.Count} filtered variables:");
            foreach (var kvp in variables)
            {
                if (kvp.Value is System.Collections.IEnumerable enumerable && !(kvp.Value is string))
                {
                    var list = enumerable.Cast<object>().ToList();
                    Logger.Debug($"  {kvp.Key} = [{string.Join(", ", list.Take(3).Select(x => x?.ToString() ?? "null"))}] (count: {list.Count})");
                }
                else
                {
                    Logger.Debug($"  {kvp.Key} = {kvp.Value}");
                }
            }

            // Apply index offset to array expressions
            if (indexOffset > 0)
            {
                dollarSignTemplate = ApplyIndexOffset(dollarSignTemplate, indexOffset);
            }

            // Configure DollarSign options
            var options = new DollarSignOptions
            {
                SupportDollarSignSyntax = true,
                ThrowOnError = false,
                UseCache = true,
                CultureInfo = System.Globalization.CultureInfo.CurrentCulture
            };

            // Evaluate the template with variables - using EvalAsync synchronously
            var result = DollarSign.EvalAsync(dollarSignTemplate, variables, options).GetAwaiter().GetResult();

            Logger.Debug($"DataBinder: Template '{dollarSignTemplate}' evaluated to '{result}'");
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"DataBinder: Error binding data to template '{template}': {ex.Message}");
            return template; // Return original template on error
        }
    }

    /// <summary>
    /// Creates context operator variables for only the expressions that are actually used (optimized version).
    /// Extracts variable names from expressions and only creates variables for those.
    /// </summary>
    private Dictionary<string, object> CreateContextOperatorVariablesFiltered(object data, ISet<string> usedExpressions)
    {
        var variables = new Dictionary<string, object>();

        if (data == null || usedExpressions == null || usedExpressions.Count == 0)
            return variables;

        // Extract variable names from the used expressions
        var requiredVariables = ExtractVariableNamesFromExpressions(usedExpressions);

        Logger.Debug($"DataBinder: Filtering variables based on {usedExpressions.Count} expressions");
        Logger.Debug($"DataBinder: Required variables: {string.Join(", ", requiredVariables)}");

        // Add the root data object
        variables["Data"] = data;

        // Create flattened variables for only the required properties
        CreateContextVariablesRecursiveFiltered(data, string.Empty, variables, 0, requiredVariables);

        return variables;
    }    /// <summary>
         /// Extracts variable names from expressions (similar to PowerPointProcessor but adapted for DataBinder).
         /// </summary>
    private HashSet<string> ExtractVariableNamesFromExpressions(ISet<string> expressions)
    {
        var variableNames = new HashSet<string>();

        foreach (var expression in expressions)
        {
            if (string.IsNullOrWhiteSpace(expression))
                continue;

            var cleanExpression = expression.Trim().TrimStart('$', '{').TrimEnd('}');

            // Handle array indexing with property access: Items[0].Name -> Items, Name
            var arrayPropertyMatch = System.Text.RegularExpressions.Regex.Match(cleanExpression, @"^([a-zA-Z_]\w*)(?:\[\d+\])\.([a-zA-Z_]\w*)");
            if (arrayPropertyMatch.Success)
            {
                variableNames.Add(arrayPropertyMatch.Groups[1].Value); // Items
                variableNames.Add(arrayPropertyMatch.Groups[2].Value); // Name (property name for filtering)
                continue;
            }

            // Handle array indexing: Items[0] -> Items
            var arrayMatch = System.Text.RegularExpressions.Regex.Match(cleanExpression, @"^([a-zA-Z_]\w*)(?:\[\d+\])?");
            if (arrayMatch.Success)
            {
                variableNames.Add(arrayMatch.Groups[1].Value);
            }

            // Handle property access: object.property -> object
            var propertyMatch = System.Text.RegularExpressions.Regex.Match(cleanExpression, @"^([a-zA-Z_]\w*)(?:\.|:)");
            if (propertyMatch.Success)
            {
                variableNames.Add(propertyMatch.Groups[1].Value);
            }

            // Handle function calls: ppt.Image(LogoPath) -> LogoPath
            var functionMatch = System.Text.RegularExpressions.Regex.Match(cleanExpression, @"ppt\.\w+\(([^)]+)\)");
            if (functionMatch.Success)
            {
                var parameters = functionMatch.Groups[1].Value.Split(',');
                foreach (var param in parameters)
                {
                    var trimmedParam = param.Trim();
                    if (!trimmedParam.StartsWith("\"") && !trimmedParam.StartsWith("'"))
                    {
                        variableNames.Add(trimmedParam);
                    }
                }
            }

            // Handle simple variable names
            if (System.Text.RegularExpressions.Regex.IsMatch(cleanExpression, @"^[a-zA-Z_]\w*$"))
            {
                variableNames.Add(cleanExpression);
            }
        }

        return variableNames;
    }

    /// <summary>
    /// Recursively creates context variables for nested objects and arrays, but only for required variables.
    /// </summary>
    private void CreateContextVariablesRecursiveFiltered(object obj, string prefix, Dictionary<string, object> variables, int depth, HashSet<string> requiredVariables)
    {
        if (obj == null || depth > 5) // Prevent infinite recursion
            return;

        // Handle Dictionary<string, object> specially
        if (obj is Dictionary<string, object> dictionary)
        {
            foreach (var kvp in dictionary)
            {
                if (kvp.Value == null)
                    continue;

                var propertyName = string.IsNullOrEmpty(prefix)
                    ? kvp.Key
                    : $"{prefix}___{kvp.Key}";

                // Only process if this variable is required
                if (IsVariableRequired(propertyName, kvp.Key, requiredVariables))
                {
                    // Add the property value
                    variables[propertyName] = kvp.Value;

                    // Handle arrays/collections - need to extract nested properties from array elements
                    if (kvp.Value is System.Collections.IEnumerable enumerable && !(kvp.Value is string))
                    {
                        CreateContextVariablesForArrayFiltered(enumerable, propertyName, variables, depth + 1, requiredVariables);

                        // For context operators, also extract common properties from array elements
                        var list = enumerable.Cast<object>().ToList();
                        if (list.Count > 0 && list[0] != null)
                        {
                            // Extract common properties from the first element to create context operator paths
                            ExtractArrayElementPropertiesFiltered(list, propertyName, variables, depth + 1, requiredVariables);
                        }
                    }
                    // Handle complex objects
                    else if (!IsSimpleType(kvp.Value.GetType()))
                    {
                        CreateContextVariablesRecursiveFiltered(kvp.Value, propertyName, variables, depth + 1, requiredVariables);
                    }
                }
            }
            return;
        }

        // Handle regular objects
        var properties = obj.GetType().GetProperties();
        foreach (var property in properties)
        {
            if (!property.CanRead)
                continue;

            try
            {
                var propertyName = string.IsNullOrEmpty(prefix)
                    ? property.Name
                    : $"{prefix}___{property.Name}";

                // Only process if this variable is required
                if (IsVariableRequired(propertyName, property.Name, requiredVariables))
                {
                    var value = property.GetValue(obj);
                    if (value == null)
                        continue;

                    variables[propertyName] = value;

                    // Handle arrays/collections - need to extract nested properties from array elements
                    if (value is System.Collections.IEnumerable enumerable && !(value is string))
                    {
                        CreateContextVariablesForArrayFiltered(enumerable, propertyName, variables, depth + 1, requiredVariables);

                        // For context operators, also extract common properties from array elements
                        var list = enumerable.Cast<object>().ToList();
                        if (list.Count > 0 && list[0] != null)
                        {
                            // Extract common properties from the first element to create context operator paths
                            ExtractArrayElementPropertiesFiltered(list, propertyName, variables, depth + 1, requiredVariables);
                        }
                    }
                    // Handle complex objects
                    else if (!IsSimpleType(value.GetType()))
                    {
                        CreateContextVariablesRecursiveFiltered(value, propertyName, variables, depth + 1, requiredVariables);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"DataBinder: Error processing property '{property.Name}': {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Checks if a variable is required based on the required variables list.
    /// </summary>
    private bool IsVariableRequired(string fullPropertyName, string propertyName, HashSet<string> requiredVariables)
    {
        return requiredVariables.Contains(propertyName) ||
               requiredVariables.Contains(fullPropertyName) ||
               requiredVariables.Any(req => req.StartsWith(propertyName + ".") || req.StartsWith(propertyName + "["));
    }    /// <summary>
         /// Creates context variables for array/collection elements (filtered version).
         /// </summary>
    private void CreateContextVariablesForArrayFiltered(System.Collections.IEnumerable enumerable, string propertyName, Dictionary<string, object> variables, int depth, HashSet<string> requiredVariables)
    {
        var list = enumerable.Cast<object>().ToList();

        // Add indexed access for each element (only if needed)
        for (int i = 0; i < list.Count; i++)
        {
            var element = list[i];
            if (element == null)
                continue; var indexedName = $"{propertyName}_{i}";

            // Only add if this indexed variable is required
            if (requiredVariables.Any(req => req.StartsWith(propertyName + "_") || req == propertyName))
            {
                variables[indexedName] = element;

                // For array elements, we need to expose their properties as flattened variables 
                // so DollarSign can access them directly (e.g., Items[0].Id -> Items_0___Id)
                if (!IsSimpleType(element.GetType()))
                {
                    var elementProperties = element.GetType().GetProperties();
                    foreach (var prop in elementProperties)
                    {
                        if (!prop.CanRead)
                            continue;

                        try
                        {
                            var flattenedName = $"{indexedName}___{prop.Name}";
                            var propValue = prop.GetValue(element);
                            if (propValue != null)
                            {
                                variables[flattenedName] = propValue;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Warning($"DataBinder: Error flattening array element property '{prop.Name}': {ex.Message}");
                        }
                    }

                    // Also recursively process for deeper nesting
                    CreateContextVariablesRecursiveFiltered(element, indexedName, variables, depth + 1, requiredVariables);
                }
            }
        }

        // Add the array itself (always needed if the property is required)
        if (requiredVariables.Contains(propertyName.Split('.').Last()) || requiredVariables.Contains(propertyName))
        {
            variables[propertyName] = list;
        }
    }

    /// <summary>
    /// Extracts properties from array elements for context operator paths (filtered version).
    /// </summary>
    private void ExtractArrayElementPropertiesFiltered(List<object> list, string arrayPropertyName, Dictionary<string, object> variables, int depth, HashSet<string> requiredVariables)
    {
        if (list.Count == 0 || list[0] == null)
            return;

        var firstElement = list[0];
        var elementProperties = firstElement.GetType().GetProperties();

        foreach (var prop in elementProperties)
        {
            if (!prop.CanRead)
                continue;

            try
            {
                var contextPath = $"{arrayPropertyName}___{prop.Name}";

                // Only process if this context path is required
                if (IsVariableRequired(contextPath, prop.Name, requiredVariables))
                {
                    var values = new List<object?>();
                    foreach (var item in list)
                    {
                        if (item != null)
                        {
                            var value = prop.GetValue(item);
                            values.Add(value);
                        }
                        else
                        {
                            values.Add(null);
                        }
                    }
                    variables[contextPath] = values;
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"DataBinder: Error extracting array element property '{prop.Name}': {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Binds data to template expressions in the given text with custom functions.
    /// Converts context operators (>) to underscore notation (___) and evaluates using DollarSignEngine.
    /// </summary>
    /// <param name="template">Template text containing expressions</param>
    /// <param name="data">Data object to bind</param>
    /// <param name="usedExpressions">Set of expressions actually used for filtering variables</param>
    /// <param name="customFunctions">Custom functions to register for use in expressions</param>
    /// <param name="indexOffset">Index offset for array expressions</param>
    /// <returns>Text with all expressions evaluated</returns>
    public string BindData(string template, object data, ISet<string> usedExpressions, Dictionary<string, Func<object, object>>? customFunctions, int indexOffset = 0)
    {
        if (string.IsNullOrEmpty(template) || data == null)
            return template ?? string.Empty;

        try
        {
            // Pre-process nested expressions first
            var preprocessedTemplate = PreprocessNestedExpressions(template, data, indexOffset);
            Logger.Debug($"DataBinder: Original template: '{template}'");
            Logger.Debug($"DataBinder: Preprocessed template: '{preprocessedTemplate}'");

            // Convert template to DollarSign format
            var dollarSignTemplate = ConvertToTemplate(preprocessedTemplate);
            Logger.Debug($"DataBinder: Converted template to '{dollarSignTemplate}'");            // Create context variables for the data - OPTIMIZED VERSION
            var variables = CreateContextOperatorVariablesFiltered(data, usedExpressions);

            // Add PPT functions to variables - reuse existing instance if available
            DocuChef.Presentation.Functions.PPTFunctions pptFunctions;
            if (data is Dictionary<string, object> dataDict &&
                dataDict.TryGetValue("ppt", out var existingPpt) &&
                existingPpt is DocuChef.Presentation.Functions.PPTFunctions existingPptFunctions)
            {
                // Reuse existing PPTFunctions instance (preserves image cache)
                pptFunctions = existingPptFunctions;
                Logger.Debug($"DataBinder: Reusing existing PPTFunctions instance with {existingPptFunctions.GetAllImageCache().Count} cached images");
            }
            else
            {
                // Create new PPTFunctions instance
                pptFunctions = new DocuChef.Presentation.Functions.PPTFunctions(variables);
                Logger.Debug($"DataBinder: Created new PPTFunctions instance");
            }
            variables["ppt"] = pptFunctions;

            // Add custom functions to variables
            if (customFunctions != null)
            {
                foreach (var kvp in customFunctions)
                {
                    variables[kvp.Key] = kvp.Value;
                }
                Logger.Debug($"DataBinder: Added {customFunctions.Count} custom functions");
            }

            Logger.Debug($"DataBinder: Created {variables.Count} filtered variables:");
            foreach (var kvp in variables)
            {
                if (kvp.Value is System.Collections.IEnumerable enumerable && !(kvp.Value is string))
                {
                    var list = enumerable.Cast<object>().ToList();
                    Logger.Debug($"  {kvp.Key} = [{string.Join(", ", list.Take(3).Select(x => x?.ToString() ?? "null"))}] (count: {list.Count})");
                }
                else
                {
                    Logger.Debug($"  {kvp.Key} = {kvp.Value}");
                }
            }

            // Apply index offset to array expressions
            if (indexOffset > 0)
            {
                dollarSignTemplate = ApplyIndexOffset(dollarSignTemplate, indexOffset);
            }

            // Configure DollarSign options
            var options = new DollarSignOptions
            {
                SupportDollarSignSyntax = true,
                ThrowOnError = false,
                UseCache = true,
                CultureInfo = System.Globalization.CultureInfo.CurrentCulture
            };

            // Evaluate the template with variables - using EvalAsync synchronously
            var result = DollarSign.EvalAsync(dollarSignTemplate, variables, options).GetAwaiter().GetResult();

            Logger.Debug($"DataBinder: Template '{dollarSignTemplate}' evaluated to '{result}'");
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"DataBinder: Error binding data to template '{template}': {ex.Message}");
            return template; // Return original template on error
        }
    }

    /// <summary>
    /// Resolves an expression using filtered variables for optimization.
    /// </summary>
    public object? ResolveExpressionWithFilteredVariables(string expression, object data, ISet<string> usedExpressions)
    {
        if (string.IsNullOrEmpty(expression) || data == null)
            return null;

        try
        {
            // Use the optimized BindData method with filtered variables
            var result = BindData(expression, data, usedExpressions);
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"DataBinder: Error resolving expression '{expression}' with filtered variables: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Resolves an expression using filtered variables and custom functions for optimization.
    /// </summary>
    public object? ResolveExpressionWithFilteredVariables(string expression, object data, ISet<string> usedExpressions, Dictionary<string, Func<object, object>>? customFunctions)
    {
        if (string.IsNullOrEmpty(expression) || data == null)
            return null;

        try
        {
            // Use the optimized BindData method with filtered variables and custom functions
            var result = BindData(expression, data, usedExpressions, customFunctions);
            return result;
        }
        catch (Exception ex)
        {
            Logger.Error($"DataBinder: Error resolving expression '{expression}' with filtered variables and custom functions: {ex.Message}");
            return null;
        }
    }
}