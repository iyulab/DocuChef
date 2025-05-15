namespace DocuChef.Helpers;

/// <summary>
/// Helper class for expression processing in document templates
/// </summary>
public static class ExpressionHelper
{
    private static readonly Regex ExpressionPattern = new Regex(@"\$\{([^{}]+)\}", RegexOptions.Compiled);

    /// <summary>
    /// Check if text contains expressions
    /// </summary>
    public static bool ContainsExpressions(string text)
    {
        return !string.IsNullOrEmpty(text) && text.Contains("${");
    }

    /// <summary>
    /// Process expressions in text using evaluator with improved handling of array indices
    /// </summary>
    public static string ProcessExpressions(string text, IExpressionEvaluator evaluator, Dictionary<string, object> variables)
    {
        if (!ContainsExpressions(text))
            return text;

        // 배열 표현식을 미리 추출하여 범위 체크
        var arrayExpressions = ExtractArrayExpressions(text);
        foreach (var expr in arrayExpressions)
        {
            // 배열 이름과 인덱스 추출
            var match = Regex.Match(expr, @"\$\{(\w+)\[(\d+)\]");
            if (match.Success && match.Groups.Count > 2)
            {
                string arrayName = match.Groups[1].Value;
                int index = int.Parse(match.Groups[2].Value);

                // 배열 변수가 있고 인덱스가 범위를 벗어나는지 확인
                if (variables.TryGetValue(arrayName, out var arrayObj) && arrayObj != null)
                {
                    int count = CollectionHelper.GetCollectionCount(arrayObj);
                    if (index >= count)
                    {
                        // 범위 초과 표현식은 빈 문자열로 대체
                        Logger.Warning($"Array index out of bounds: {arrayName}[{index}] exceeds {count} items");
                        text = text.Replace(expr, "");
                    }
                }
            }
        }

        // 표준 표현식 처리 
        return ExpressionPattern.Replace(text, match =>
        {
            try
            {
                var result = evaluator.EvaluateCompleteExpression(match.Value, variables);
                // null 결과는 빈 문자열로 처리
                return result?.ToString() ?? "";
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error evaluating expression '{match.Value}': {ex.Message}");
                return ""; // 오류 발생 시 표현식을 빈 문자열로 대체
            }
        });
    }

    /// <summary>
    /// Extract array expressions from text
    /// </summary>
    private static List<string> ExtractArrayExpressions(string text)
    {
        var result = new List<string>();
        var pattern = new Regex(@"\$\{(\w+)\[(\d+)\][^}]*\}");

        var matches = pattern.Matches(text);
        foreach (Match match in matches)
        {
            result.Add(match.Value);
        }

        return result;
    }

    /// <summary>
    /// Check if the entire text is a single expression
    /// </summary>
    public static bool IsSingleExpression(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        text = text.Trim();
        return text.StartsWith("${") && text.EndsWith("}") &&
               text.IndexOf("${", 2) == -1;
    }

    /// <summary>
    /// Parse function parameters string into array
    /// </summary>
    public static string[] ParseFunctionParameters(string parametersString)
    {
        if (string.IsNullOrEmpty(parametersString))
            return Array.Empty<string>();

        var results = new List<string>();
        bool inQuotes = false;
        int currentStart = 0;

        for (int i = 0; i < parametersString.Length; i++)
        {
            char c = parametersString[i];

            // Handle quotes
            if (c == '"' && (i == 0 || parametersString[i - 1] != '\\'))
            {
                inQuotes = !inQuotes;
            }
            // Handle parameter separators
            else if (c == ',' && !inQuotes)
            {
                results.Add(parametersString.Substring(currentStart, i - currentStart).Trim());
                currentStart = i + 1;
            }
        }

        // Add the last parameter
        if (currentStart < parametersString.Length)
        {
            results.Add(parametersString.Substring(currentStart).Trim());
        }

        // Clean up parameters
        return results.Select(CleanParameter).ToArray();
    }

    /// <summary>
    /// Clean up a parameter value
    /// </summary>
    private static string CleanParameter(string param)
    {
        param = param.Trim();

        // Remove surrounding quotes
        if (param.StartsWith("\"") && param.EndsWith("\"") && param.Length > 1)
        {
            param = param.Substring(1, param.Length - 2)
                .Replace("\\\"", "\"")
                .Replace("\\\\", "\\");
        }

        return param;
    }
}