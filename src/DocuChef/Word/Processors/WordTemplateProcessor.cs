using DocuChef.Word.Functions;
using DocuChef.Word.Models;
using WBody = DocumentFormat.OpenXml.Wordprocessing.Body;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;
using WRun = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace DocuChef.Word.Processors;

/// <summary>
/// Orchestrates the 7-stage Word template processing pipeline:
/// 1. PreprocessRuns — merge split expression runs
/// 2. ProcessRepeaters — expand table rows and paragraph blocks
/// 3. Re-merge — merge runs again after repeater expansion
/// 4. ExtractImagePlaceholders — find and parse word.Image() expressions
/// 5. BindText — evaluate ${expression} placeholders
/// 6. ProcessImages — replace image markers with Drawing elements
/// 7. Cleanup — remove unresolved expressions
/// </summary>
public class WordTemplateProcessor
{
    private static readonly Regex ImageExpressionPattern =
        new(@"\$\{word\.Image\(([^)]+)\)\}", RegexOptions.Compiled);

    /// <summary>
    /// Processes the template document and returns the result as a MemoryStream.
    /// The template document is opened read-only; a working copy is created in memory.
    /// </summary>
    public MemoryStream Process(
        WordprocessingDocument templateDoc,
        WordOptions options,
        Dictionary<string, object> data)
    {
        // Clone template to a writable memory stream
        var outputStream = CloneToMemory(templateDoc);

        using var workingDoc = WordprocessingDocument.Open(outputStream, true);
        var mainPart = workingDoc.MainDocumentPart
            ?? throw new DocuChefException("Template has no main document part.");
        var document = mainPart.Document
            ?? throw new DocuChefException("Template has no document element.");
        var body = document.Body
            ?? throw new DocuChefException("Template document has no body.");

        // Stage 1: Preprocess runs — merge split expressions
        Logger.Debug("Pipeline stage 1: PreprocessRuns");
        RunMerger.MergeExpressionRuns(body);

        // Stage 2: Process repeaters — expand table rows and paragraph blocks
        Logger.Debug("Pipeline stage 2: ProcessRepeaters");
        TableRepeater.ProcessTables(body, data);
        ParagraphRepeater.ProcessParagraphs(body, data);

        // Stage 3: Re-merge — new paragraphs from repeaters may have split runs
        Logger.Debug("Pipeline stage 3: Re-merge runs");
        RunMerger.MergeExpressionRuns(body);

        // Stage 4: Extract image placeholders before text binding
        Logger.Debug("Pipeline stage 4: ExtractImagePlaceholders");
        var imagePlaceholders = ExtractImagePlaceholders(body, data);

        // Stage 5: Bind text — evaluate ${expression} placeholders
        Logger.Debug("Pipeline stage 5: BindText");
        TextBinder.Bind(body, data);

        // Stage 6: Process images — replace markers with Drawing elements
        Logger.Debug("Pipeline stage 6: ProcessImages");
        if (imagePlaceholders.Count > 0)
        {
            ImageHandler.ProcessImages(mainPart, body, imagePlaceholders);
        }

        // Stage 7: Cleanup — optionally remove unresolved expressions
        Logger.Debug("Pipeline stage 7: Cleanup");
        CleanupUnresolvedExpressions(body);

        // Save changes
        mainPart.Document.Save();

        return outputStream;
    }

    /// <summary>
    /// Clones the template document into a writable MemoryStream.
    /// </summary>
    private static MemoryStream CloneToMemory(WordprocessingDocument templateDoc)
    {
        var stream = new MemoryStream();
        using (var clone = templateDoc.Clone(stream))
        {
            // Clone is saved and disposed; stream now contains the copy
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Finds ${word.Image(...)} expressions, evaluates them via WordFunctions,
    /// and replaces each with a unique marker key. Returns a dictionary mapping
    /// marker keys to ImagePlaceholder instances.
    /// </summary>
    private static Dictionary<string, ImagePlaceholder> ExtractImagePlaceholders(
        WBody body, Dictionary<string, object> data)
    {
        var placeholders = new Dictionary<string, ImagePlaceholder>();

        if (!data.TryGetValue("word", out var wordObj) || wordObj is not WordFunctions wordFunctions)
            return placeholders;

        int counter = 0;
        foreach (var run in body.Descendants<WRun>().ToList())
        {
            var textElement = run.GetFirstChild<WText>();
            if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                continue;

            var text = textElement.Text;
            if (!text.Contains("word.Image"))
                continue;

            textElement.Text = ImageExpressionPattern.Replace(text, match =>
            {
                var argsString = match.Groups[1].Value;
                var placeholder = ParseImageArgs(argsString, wordFunctions);
                if (placeholder == null)
                    return match.Value; // Leave unresolved

                var markerKey = $"__DOCUCHEF_IMG_{counter++}__";
                placeholders[markerKey] = placeholder;
                return markerKey;
            });
        }

        return placeholders;
    }

    /// <summary>
    /// Parses the arguments from a word.Image() call.
    /// Supports: Image("path"), Image("path", width), Image("path", width, height)
    /// </summary>
    private static ImagePlaceholder? ParseImageArgs(string argsString, WordFunctions wordFunctions)
    {
        try
        {
            var args = SplitArgs(argsString);
            if (args.Count == 0)
                return null;

            var path = args[0].Trim().Trim('"', '\'');
            int? width = args.Count > 1 ? int.Parse(args[1].Trim()) : null;
            int? height = args.Count > 2 ? int.Parse(args[2].Trim()) : null;

            return wordFunctions.Image(path, width, height);
        }
        catch (Exception ex)
        {
            Logger.Debug($"WordTemplateProcessor: Failed to parse image args '{argsString}': {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Splits comma-separated arguments, respecting quoted strings.
    /// </summary>
    private static List<string> SplitArgs(string input)
    {
        var args = new List<string>();
        var current = new System.Text.StringBuilder();
        bool inQuote = false;
        char quoteChar = '"';

        foreach (var ch in input)
        {
            if (!inQuote && (ch == '"' || ch == '\''))
            {
                inQuote = true;
                quoteChar = ch;
                current.Append(ch);
            }
            else if (inQuote && ch == quoteChar)
            {
                inQuote = false;
                current.Append(ch);
            }
            else if (!inQuote && ch == ',')
            {
                args.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(ch);
            }
        }

        if (current.Length > 0)
            args.Add(current.ToString());

        return args;
    }

    /// <summary>
    /// Removes or clears any remaining ${...} expressions that were not resolved.
    /// </summary>
    private static void CleanupUnresolvedExpressions(WBody body)
    {
        var unresolvedPattern = new Regex(@"\$\{[^}]+\}");

        foreach (var run in body.Descendants<WRun>().ToList())
        {
            var textElement = run.GetFirstChild<WText>();
            if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                continue;

            if (unresolvedPattern.IsMatch(textElement.Text))
            {
                Logger.Debug($"Cleanup: removing unresolved expression in '{textElement.Text}'");
                textElement.Text = unresolvedPattern.Replace(textElement.Text, string.Empty);
            }
        }
    }
}
