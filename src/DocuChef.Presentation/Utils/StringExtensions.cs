namespace DocuChef.Presentation.Utils;

public static class StringExtensions
{
    public static string? GetBetween(this string input, string start, string end)
    {
        if (string.IsNullOrEmpty(input))
            return null;
        int startIndex = input.IndexOf(start, StringComparison.Ordinal);
        if (startIndex < 0)
            return null;
        startIndex += start.Length;
        int endIndex = input.IndexOf(end, startIndex, StringComparison.Ordinal);
        if (endIndex < 0)
            return null;
        return input[startIndex..endIndex];
    }
}
