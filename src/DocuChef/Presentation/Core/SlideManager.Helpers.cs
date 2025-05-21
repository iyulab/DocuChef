namespace DocuChef.Presentation.Core;

/// <summary>
/// Common slide-related utilities for use across SlideManager partial classes
/// </summary>
public static partial class SlideManager
{
    /// <summary>
    /// Checks if a text element contains a complete expression (${...})
    /// </summary>
    private static bool ContainsCompleteExpression(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        // Check for ${...} pattern
        int openIndex = text.IndexOf("${");
        if (openIndex >= 0)
        {
            int closeIndex = text.IndexOf("}", openIndex);
            if (closeIndex > openIndex)
            {
                // Make sure there's no other ${ before the closing }
                string between = text.Substring(openIndex + 2, closeIndex - openIndex - 2);
                return !between.Contains("${");
            }
        }

        // Check for $..$ pattern
        openIndex = text.IndexOf("$");
        if (openIndex >= 0 && openIndex < text.Length - 1)
        {
            int closeIndex = text.IndexOf("$", openIndex + 1);
            if (closeIndex > openIndex)
            {
                // Make sure there's no other $ in between
                string between = text.Substring(openIndex + 1, closeIndex - openIndex - 1);
                return !between.Contains("$");
            }
        }

        return false;
    }

    /// <summary>
    /// Finds all expression ranges in a text string
    /// </summary>
    internal static List<ExpressionRange> FindExpressionRanges(string text)
    {
        var result = new List<ExpressionRange>();
        int pos = 0;

        while (pos < text.Length)
        {
            int startPos = text.IndexOf("${", pos);
            if (startPos == -1)
                break;

            int endPos = text.IndexOf("}", startPos);
            if (endPos == -1)
                break;

            result.Add(new ExpressionRange
            {
                Start = startPos,
                End = endPos,
                Expression = text.Substring(startPos, endPos - startPos + 1)
            });

            pos = endPos + 1;
        }

        return result;
    }

    /// <summary>
    /// Gets the next available shape ID from the shape tree
    /// </summary>
    private static uint GetNextShapeId(P.ShapeTree shapeTree)
    {
        uint maxId = 0;

        // Find the highest Id in use
        foreach (var element in shapeTree.Descendants<P.NonVisualDrawingProperties>())
        {
            if (element.Id != null && element.Id.Value > maxId)
            {
                maxId = element.Id.Value;
            }
        }

        // Start with ID 2 at minimum (1 is typically used by the shapeTree itself)
        return Math.Max(maxId + 1, 2);
    }

    /// <summary>
    /// Checks if the specified text is a directive
    /// </summary>
    private static bool IsDirective(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        text = text.Trim();

        // Check for directive markers
        return text.StartsWith("#foreach") ||
               text.StartsWith("#if") ||
               text.Contains("#foreach:") ||
               text.Contains("#if:");
    }

    /// <summary>
    /// Updates image references in the slide
    /// </summary>
    private static void UpdateImageReferences(P.Slide slide, string oldId, string newId)
    {
        // Update blip elements (images)
        foreach (var blip in slide.Descendants<D.Blip>())
        {
            if (blip.Embed != null && blip.Embed.Value == oldId)
            {
                blip.Embed.Value = newId;
            }
        }
    }

    /// <summary>
    /// Creates a new notes slide with proper OpenXML structure
    /// </summary>
    private static P.NotesSlide CreateNewNotesSlide()
    {
        // Create the shape tree with required elements according to OpenXML spec
        var shapeTree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties() { Id = 1U, Name = "Notes" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            new P.GroupShapeProperties(
                new D.TransformGroup(
                    new D.Offset() { X = 0L, Y = 0L },
                    new D.Extents() { Cx = 0L, Cy = 0L },
                    new D.ChildOffset() { X = 0L, Y = 0L },
                    new D.ChildExtents() { Cx = 0L, Cy = 0L }
                )
            )
        );

        // Create the notes slide with required elements
        var notesSlide = new P.NotesSlide(
            new P.CommonSlideData(shapeTree)
        );

        // Add color mapping
        notesSlide.AppendChild(new P.ColorMapOverride(new D.MasterColorMapping()));

        return notesSlide;
    }

    /// <summary>
    /// Creates a TextBody element with the specified content with proper OpenXML structure
    /// </summary>
    private static P.TextBody CreateTextBody(string content)
    {
        var textBody = new P.TextBody();

        // Add required properties to TextBody based on standard PowerPoint defaults
        textBody.AppendChild(new D.BodyProperties() { Anchor = D.TextAnchoringTypeValues.Top });
        textBody.AppendChild(new D.ListStyle());

        // Create paragraph with text
        var paragraph = new D.Paragraph();
        var run = new D.Run();
        var text = new D.Text() { Text = content };

        run.AppendChild(text);
        paragraph.AppendChild(run);
        textBody.AppendChild(paragraph);

        return textBody;
    }

    /// <summary>
    /// Helper class for tracking expression ranges
    /// </summary>
    public class ExpressionRange
    {
        public int Start { get; set; }
        public int End { get; set; }
        public string Expression { get; set; }
    }

    /// <summary>
    /// Helper class for tracking text segments
    /// </summary>
    public class TextSegment
    {
        public int Start { get; set; }
        public int End { get; set; }
        public string Text { get; set; }
        public bool IsExpression { get; set; }
    }

    /// <summary>
    /// Helper class for tracking run information
    /// </summary>
    internal class RunInfo
    {
        public D.Text TextElement { get; set; }
        public int StartPosition { get; set; }
        public int EndPosition { get; set; }
        public D.Run Run { get; set; }
    }
}