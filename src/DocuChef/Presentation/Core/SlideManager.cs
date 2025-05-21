using DocuChef.Presentation.Models;

namespace DocuChef.Presentation.Core;

/// <summary>
/// Provides slide management functionality
/// </summary>
public static partial class SlideManager
{
    /// <summary>
    /// Gets the text of the slide note
    /// </summary>
    public static string GetSlideNoteText(SlidePart slidePart)
    {
        if (slidePart == null)
            return string.Empty;

        try
        {
            // Using NotesSlidePart if available
            if (slidePart.NotesSlidePart != null)
            {
                var notesSlide = slidePart.NotesSlidePart.NotesSlide;

                // Extract all text elements from notes slide
                var textElements = notesSlide.Descendants<D.Text>().ToList();

                // Log the number of text elements found
                Logger.Debug($"Found {textElements.Count} text elements in slide notes");

                // Concatenate all text content for logging
                for (int i = 0; i < textElements.Count; i++)
                {
                    var text = textElements[i];
                    Logger.Debug($"Text element {i}: '{text.Text}'");
                }

                // HANDLE MULTIPLE TEXT ELEMENTS: PowerPoint may split text into multiple elements
                // First check if we need to combine multiple text elements to form a complete directive
                if (textElements.Count > 1)
                {
                    // Special handling for combined directives when multiple text elements exist
                    string combinedText = CombineDirectiveTextElements(textElements);
                    if (!string.IsNullOrEmpty(combinedText))
                    {
                        Logger.Debug($"Combined directive text: '{combinedText}'");
                        return combinedText;
                    }
                }

                // Standard processing for single text elements
                // First look for text that might contain directives
                foreach (var text in textElements)
                {
                    if (!string.IsNullOrEmpty(text.Text) && text.Text.Trim().StartsWith("#"))
                    {
                        // This looks like a directive
                        Logger.Debug($"Found potential directive: '{text.Text}'");
                        return text.Text;
                    }
                }

                // If no potential directive found but there's text, return the first non-empty one
                foreach (var text in textElements)
                {
                    if (!string.IsNullOrEmpty(text.Text))
                    {
                        Logger.Debug($"No directive found, returning first text: '{text.Text}'");
                        return text.Text;
                    }
                }
            }
            else
            {
                Logger.Debug("No notes slide part found");
            }

            // Fallback - look for text in slide elements that might be directives
            var allSlideTexts = slidePart.Slide.Descendants<D.Text>().ToList();
            Logger.Debug($"Fallback: Found {allSlideTexts.Count} text elements in main slide");

            foreach (var text in allSlideTexts)
            {
                if (!string.IsNullOrEmpty(text.Text) && text.Text.Trim().StartsWith("#"))
                {
                    Logger.Debug($"Found potential directive in slide: '{text.Text}'");
                    return text.Text;
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error getting slide note text: {ex.Message}", ex);
        }

        Logger.Debug("No text found that could be a directive");
        return string.Empty;
    }

    /// <summary>
    /// Combines multiple text elements that might form a directive
    /// This is necessary because PowerPoint may split text across multiple runs or paragraphs
    /// </summary>
    private static string CombineDirectiveTextElements(List<D.Text> textElements)
    {
        // First find a directive marker (text starting with #)
        int directiveStartIndex = -1;

        for (int i = 0; i < textElements.Count; i++)
        {
            string text = textElements[i].Text?.Trim() ?? "";
            if (text.StartsWith("#"))
            {
                directiveStartIndex = i;
                break;
            }
        }

        if (directiveStartIndex < 0)
            return string.Empty;

        // We found a directive start, now combine with subsequent elements
        var combinedText = new StringBuilder();

        // Add the directive start
        combinedText.Append(textElements[directiveStartIndex].Text.Trim());

        // Check if this is a directive that needs a collection name (like foreach or if)
        bool isDirectiveNeedingParams = textElements[directiveStartIndex].Text.Contains("foreach") ||
                                       textElements[directiveStartIndex].Text.Contains("if");

        // If the directive text ends with a colon but no content, we need to append the next element
        bool needsMoreContent = combinedText.ToString().TrimEnd().EndsWith(":") ||
                               isDirectiveNeedingParams;

        // If we need more content and there are more elements, add them
        if (needsMoreContent && directiveStartIndex < textElements.Count - 1)
        {
            // Add a space if needed, then the next element's text
            if (!combinedText.ToString().TrimEnd().EndsWith(":"))
                combinedText.Append(" ");

            combinedText.Append(textElements[directiveStartIndex + 1].Text.Trim());

            // If there's a third element that might be part of the directive (like max items), add it too
            if (directiveStartIndex + 2 < textElements.Count &&
                !string.IsNullOrWhiteSpace(textElements[directiveStartIndex + 2].Text) &&
                textElements[directiveStartIndex + 2].Text.Trim().Length < 10) // Only add short text elements
            {
                combinedText.Append(" ");
                combinedText.Append(textElements[directiveStartIndex + 2].Text.Trim());
            }
        }

        return combinedText.ToString().Trim();
    }

    /// <summary>
    /// Updates a slide note with context information, only if the context contains a directive
    /// </summary>
    public static void UpdateSlideNote(SlidePart slidePart, SlideContext context)
    {
        if (slidePart == null || context == null)
            return;

        try
        {
            // Get context description
            string contextDescription = context.GetContextDescription();

            // Skip if not a directive - this is the key change to prevent unnecessary note creation
            if (!IsDirective(contextDescription))
            {
                Logger.Debug($"Skipping note update as context is not a directive: {contextDescription}");
                return;
            }

            // Only update existing notes slide part if available
            NotesSlidePart notesSlidePart = slidePart.NotesSlidePart;

            // If notes slide exists, update it
            if (notesSlidePart != null && notesSlidePart.NotesSlide != null)
            {
                // Update note content
                UpdateNoteContent(notesSlidePart.NotesSlide, contextDescription);
                Logger.Debug($"Updated existing slide note: {contextDescription}");
            }
            else
            {
                // Don't create new notes for non-directive contexts
                Logger.Debug($"Notes slide part not found, and we're not creating one for non-directive contexts");
            }
        }
        catch (Exception ex)
        {
            Logger.Debug($"Error updating slide note: {ex.Message}");
        }
    }

    /// <summary>
    /// Updates a slide note with an implicit directive
    /// </summary>
    public static void UpdateSlideNoteWithImplicitDirective(SlidePart slidePart, string directiveText)
    {
        if (slidePart == null || string.IsNullOrEmpty(directiveText))
            return;

        // Skip if not a directive
        if (!IsDirective(directiveText))
        {
            Logger.Debug($"Skipping note update as text is not a directive: {directiveText}");
            return;
        }

        try
        {
            // Get existing note text
            string existingNote = GetSlideNoteText(slidePart);

            // Only add the directive if it doesn't already exist
            if (string.IsNullOrEmpty(existingNote) || !existingNote.Contains(directiveText))
            {
                // Get or create notes slide part
                NotesSlidePart notesSlidePart = slidePart.NotesSlidePart;

                if (notesSlidePart == null)
                {
                    // Create notes slide if it doesn't exist - this is a directive so it's important
                    notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
                    notesSlidePart.NotesSlide = CreateNewNotesSlide();
                }

                string newNoteText = directiveText;
                if (!string.IsNullOrEmpty(existingNote) && !existingNote.StartsWith("#"))
                {
                    // Preserve existing note text if it's not a directive
                    newNoteText = $"{directiveText}\n{existingNote}";
                }

                // Update note content
                UpdateNoteContent(notesSlidePart.NotesSlide, newNoteText);
                Logger.Debug($"Added implicit directive to slide note: {directiveText}");
            }
            else
            {
                Logger.Debug($"Slide note already contains directive, no update needed");
            }
        }
        catch (Exception ex)
        {
            Logger.Debug($"Error updating slide note with implicit directive: {ex.Message}");
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
    /// Updates the note content with proper OpenXML structure
    /// </summary>
    private static void UpdateNoteContent(P.NotesSlide notesSlide, string content)
    {
        // Make sure we have a valid notes slide and shape tree
        if (notesSlide == null)
            return;

        // Find shape tree
        P.ShapeTree shapeTree = notesSlide.Descendants<P.ShapeTree>().FirstOrDefault();
        if (shapeTree == null)
        {
            Logger.Warning("Cannot update note content: ShapeTree not found");
            return;
        }

        try
        {
            // Find existing note text field
            var existingShape = shapeTree.Elements<P.Shape>()
                .FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Notes Text");

            if (existingShape != null)
            {
                // Update existing text
                var existingTextBody = existingShape.Elements<P.TextBody>().FirstOrDefault();
                if (existingTextBody != null)
                {
                    UpdateTextBodyContent(existingTextBody, content);
                }
                else
                {
                    // Create text body if it doesn't exist
                    existingShape.AppendChild(CreateTextBody(content));
                }
            }
            else
            {
                // Create new shape for notes with standard layout based on PowerPoint OpenXML
                uint nextId = GetNextShapeId(shapeTree);

                // Create a shape based on standard PowerPoint notes layout
                P.Shape shape = new P.Shape();

                // Non-visual properties with standard PowerPoint notes metadata
                shape.AppendChild(new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = nextId, Name = "Notes Text" },
                    new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                    new P.ApplicationNonVisualDrawingProperties()
                ));

                // Shape properties based on standard PowerPoint notes layout
                shape.AppendChild(new P.ShapeProperties(
                    new D.Transform2D(
                        new D.Offset() { X = 3600000L, Y = 1800000L },
                        new D.Extents() { Cx = 5400000L, Cy = 3600000L }
                    ),
                    new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }
                ));

                // Create text body with content
                shape.AppendChild(CreateTextBody(content));

                // Add shape to shape tree
                shapeTree.AppendChild(shape);
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error updating note content: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Updates the content of a TextBody element, preserving required structure
    /// </summary>
    private static void UpdateTextBodyContent(P.TextBody textBody, string content)
    {
        // Make sure body properties exist, based on standard PowerPoint defaults
        if (textBody.BodyProperties == null)
        {
            textBody.AppendChild(new D.BodyProperties());
        }

        // Make sure list style exists
        if (textBody.ListStyle == null)
        {
            textBody.AppendChild(new D.ListStyle());
        }

        // Clear existing paragraphs
        var paragraphs = textBody.Elements<D.Paragraph>().ToList();
        foreach (var para in paragraphs)
        {
            para.Remove();
        }

        // Create and add new paragraph
        var newParagraph = new D.Paragraph();
        var newRun = new D.Run();
        var newText = new D.Text() { Text = content };

        newRun.AppendChild(newText);
        newParagraph.AppendChild(newRun);
        textBody.AppendChild(newParagraph);
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
}