using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing.ArrayProcessing;

/// <summary>
/// Processes array-based batch operations for PowerPoint slides
/// </summary>
internal partial class ArrayBatchProcessor
{
    private readonly PowerPointContext _context;
    private readonly Dictionary<string, object> _variables;

    /// <summary>
    /// Creates a new instance of the ArrayBatchProcessor
    /// </summary>
    public ArrayBatchProcessor(PowerPointContext context, Dictionary<string, object> variables)
    {
        _context = context ?? throw new ArgumentNullException(nameof(context));
        _variables = variables ?? throw new ArgumentNullException(nameof(variables));
    }

    /// <summary>
    /// Auto-detect maximum items per slide based on array references
    /// </summary>
    private int DetectMaxItemsPerSlide(SlidePart slidePart, string arrayName)
    {
        var arrayReferences = FindArrayReferencesInSlide(slidePart)
            .Where(r => r.ArrayName == arrayName)
            .ToList();

        if (!arrayReferences.Any())
            return 1; // Default to 1 if no references found

        // Find the maximum index referenced
        int maxIndex = arrayReferences.Max(r => r.Index);

        // Max items per slide is max index + 1
        return maxIndex + 1;
    }

    /// <summary>
    /// Find all array references in slide
    /// </summary>
    private List<ArrayReference> FindArrayReferencesInSlide(SlidePart slidePart)
    {
        var result = new List<ArrayReference>();

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (shape.TextBody == null)
                continue;

            var references = PowerPointShapeHelper.FindArrayReferences(shape);
            result.AddRange(references);
        }

        return result;
    }

    /// <summary>
    /// Update array references in slide with specified offset
    /// </summary>
    private void UpdateArrayReferences(SlidePart slidePart, string arrayName, int offset)
    {
        Logger.Debug($"Updating array references: {arrayName} with offset {offset}");

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (shape.TextBody == null)
                continue;

            foreach (var text in shape.Descendants<A.Text>())
            {
                if (string.IsNullOrEmpty(text.Text) || !text.Text.Contains(arrayName))
                    continue;

                string original = text.Text;

                // 1. Process ${ArrayName[n]} pattern
                var dollarPattern = new Regex($@"\$\{{{arrayName}\[(\d+)\]([^\}}]*)\}}");
                text.Text = dollarPattern.Replace(text.Text, match => {
                    int index = int.Parse(match.Groups[1].Value);
                    string suffix = match.Groups[2].Value;
                    return $"${{{arrayName}[{index + offset}]{suffix}}}";
                });

                // 2. Process direct ArrayName[n] pattern
                var directPattern = new Regex($@"(?<!\$\{{){arrayName}\[(\d+)\]");
                text.Text = directPattern.Replace(text.Text, match => {
                    int index = int.Parse(match.Groups[1].Value);
                    return $"{arrayName}[{index + offset}]";
                });

                if (text.Text != original)
                {
                    Logger.Debug($"Updated text from '{original}' to '{text.Text}'");
                }
            }
        }
    }

    /// <summary>
    /// Hide shapes with out-of-range array indices
    /// </summary>
    private void HideOutOfRangeItems(SlidePart slidePart, string arrayName, int startIndex, int totalItems, int maxItems)
    {
        Logger.Debug($"Hiding out-of-range shapes: {arrayName}, startIndex={startIndex}, totalItems={totalItems}, maxItems={maxItems}");

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (shape.TextBody == null || PowerPointShapeHelper.IsShapeHidden(shape))
                continue;

            bool hasOutOfRangeReference = false;
            foreach (var text in shape.Descendants<A.Text>())
            {
                if (string.IsNullOrEmpty(text.Text))
                    continue;

                // Check for array references
                var matches = Regex.Matches(text.Text, $@"\$\{{{arrayName}\[(\d+)\]");
                foreach (Match match in matches)
                {
                    if (match.Groups.Count < 2)
                        continue;

                    int referencedIndex = int.Parse(match.Groups[1].Value);

                    // Check if referenced index is out of range or beyond this batch's range
                    if (referencedIndex >= totalItems)
                    {
                        hasOutOfRangeReference = true;
                        Logger.Debug($"Found out-of-range reference: {arrayName}[{referencedIndex}] >= {totalItems}");
                        break;
                    }

                    // Check batch range
                    int batchEndIndex = startIndex + maxItems - 1;
                    if (referencedIndex < startIndex || referencedIndex > batchEndIndex)
                    {
                        hasOutOfRangeReference = true;
                        Logger.Debug($"Found out-of-batch reference: {arrayName}[{referencedIndex}] outside range {startIndex}-{batchEndIndex}");
                        break;
                    }
                }

                if (hasOutOfRangeReference)
                    break;
            }

            if (hasOutOfRangeReference)
            {
                // Hide the shape
                PowerPointShapeHelper.HideShape(shape);
                Logger.Debug($"Hidden shape with out-of-range reference to {arrayName}");
            }
        }
    }

    /// <summary>
    /// Remove a slide from presentation
    /// </summary>
    private void RemoveSlide(PresentationPart presentationPart, SlidePart slidePart)
    {
        var slideIds = presentationPart.Presentation.SlideIdList;
        string relationshipId = presentationPart.GetIdOfPart(slidePart);
        var slideIdToRemove = slideIds.Elements<SlideId>()
            .FirstOrDefault(s => s.RelationshipId == relationshipId);

        if (slideIdToRemove != null)
        {
            slideIds.RemoveChild(slideIdToRemove);
            Logger.Debug($"Removed slide from presentation");
        }
    }
}