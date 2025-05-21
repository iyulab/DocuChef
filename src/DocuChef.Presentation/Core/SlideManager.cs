namespace DocuChef.Presentation.Core;

/// <summary>
/// Provides slide management functionality
/// </summary>
public static class SlideManager
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

                // First look for text that might contain directives
                // Instead of hardcoding "#foreach" or "#if", we'll look for text 
                // starting with "#" which is more general
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
    /// Updates a slide note with context information
    /// </summary>
    public static void UpdateSlideNote(SlidePart slidePart, Models.SlideContext context)
    {
        if (slidePart == null || context == null)
            return;

        try
        {
            // Get notes slide part
            NotesSlidePart notesSlidePart = slidePart.NotesSlidePart;

            if (notesSlidePart == null)
            {
                // Create notes slide if it doesn't exist
                notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
                notesSlidePart.NotesSlide = CreateNewNotesSlide();
            }

            // Get context description
            string contextDescription = context.GetContextDescription();

            // Update note content
            UpdateNoteContent(notesSlidePart.NotesSlide, contextDescription);

            Logger.Debug($"Updated slide note: {contextDescription}");
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

        try
        {
            // Get notes slide part
            NotesSlidePart notesSlidePart = slidePart.NotesSlidePart;

            if (notesSlidePart == null)
            {
                // Create notes slide if it doesn't exist
                notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
                notesSlidePart.NotesSlide = CreateNewNotesSlide();
            }

            // Get existing note text
            string existingNote = GetSlideNoteText(slidePart);

            // Only add the directive if it doesn't already exist
            if (string.IsNullOrEmpty(existingNote) || !existingNote.Contains(directiveText))
            {
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
    /// Creates a new notes slide
    /// </summary>
    private static P.NotesSlide CreateNewNotesSlide()
    {
        var notesSlide = new P.NotesSlide(new P.CommonSlideData(new P.ShapeTree()));
        notesSlide.AppendChild(new P.ColorMapOverride(new D.MasterColorMapping()));

        return notesSlide;
    }

    /// <summary>
    /// Updates the note content
    /// </summary>
    private static void UpdateNoteContent(P.NotesSlide notesSlide, string content)
    {
        // Find or create shape tree
        P.ShapeTree shapeTree = notesSlide.Descendants<P.ShapeTree>().FirstOrDefault();
        if (shapeTree == null)
        {
            // Create ShapeTree if it doesn't exist
            P.CommonSlideData commonSlideData = notesSlide.CommonSlideData ??
                notesSlide.AppendChild(new P.CommonSlideData());

            if (commonSlideData.ShapeTree == null)
            {
                shapeTree = commonSlideData.AppendChild(new P.ShapeTree());

                // Add required children to ShapeTree
                shapeTree.AppendChild(new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = 1U, Name = "Notes" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()));

                shapeTree.AppendChild(new P.GroupShapeProperties(new D.TransformGroup()));
            }
            else
            {
                shapeTree = commonSlideData.ShapeTree;
            }
        }

        // Find existing note text field
        var existingShape = shapeTree.Elements<P.Shape>()
            .FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Context Note");

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
            // Create new shape
            uint maxId = GetMaxShapeId(shapeTree);
            uint newId = maxId + 1;

            // Create shape with all required properties
            P.Shape shape = new P.Shape();

            // Non-visual properties
            shape.AppendChild(new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = newId, Name = "Context Note" },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()));

            // Shape properties
            shape.AppendChild(new P.ShapeProperties(
                new D.Transform2D(
                    new D.Offset() { X = 1270000L, Y = 1270000L },
                    new D.Extents() { Cx = 6096000L, Cy = 1483200L }
                ),
                new D.PresetGeometry(new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }
            ));

            // Create text body with content
            shape.AppendChild(CreateTextBody(content));

            // Add shape to shape tree
            shapeTree.AppendChild(shape);
        }
    }

    /// <summary>
    /// Updates the content of a TextBody element
    /// </summary>
    private static void UpdateTextBodyContent(P.TextBody textBody, string content)
    {
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
    /// Creates a TextBody element with the specified content
    /// </summary>
    private static P.TextBody CreateTextBody(string content)
    {
        var textBody = new P.TextBody();

        // Add required properties to TextBody
        textBody.AppendChild(new D.BodyProperties());
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
    /// Gets the maximum shape ID from the shape tree
    /// </summary>
    private static uint GetMaxShapeId(P.ShapeTree shapeTree)
    {
        uint maxId = 1;

        // Find the highest Id in use
        foreach (var element in shapeTree.Descendants<P.NonVisualDrawingProperties>())
        {
            if (element.Id != null && element.Id.Value > maxId)
            {
                maxId = element.Id.Value;
            }
        }

        return maxId;
    }

    /// <summary>
    /// Clones a slide with all its relationships
    /// </summary>
    public static SlidePart CloneSlide(PresentationPart presentationPart, SlidePart sourceSlidePart)
    {
        if (presentationPart == null || sourceSlidePart == null)
            throw new ArgumentNullException("Source or presentation part is null");

        // Add a new slide part
        SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>();

        // Clone slide content
        newSlidePart.Slide = (P.Slide)sourceSlidePart.Slide.CloneNode(true);

        // Clone slide layout relationship
        if (sourceSlidePart.SlideLayoutPart != null)
        {
            newSlidePart.AddPart(sourceSlidePart.SlideLayoutPart);
        }

        // Clone notes relationship and content if present
        if (sourceSlidePart.NotesSlidePart != null)
        {
            CloneNotesPart(sourceSlidePart, newSlidePart);
        }

        // Clone other important relationships (images, charts, etc.)
        CloneSlideRelationships(sourceSlidePart, newSlidePart);

        return newSlidePart;
    }

    /// <summary>
    /// Clones a slide and adjusts expression indices based on slide context
    /// </summary>
    public static SlidePart CloneSlideWithContext(PresentationPart presentationPart, SlidePart sourceSlidePart, Models.SlideContext context)
    {
        if (presentationPart == null || sourceSlidePart == null)
            throw new ArgumentNullException("Source or presentation part is null");

        // Add a new slide part
        SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>();

        // Clone slide content
        newSlidePart.Slide = (P.Slide)sourceSlidePart.Slide.CloneNode(true);

        // Adjust expressions in text elements
        AdjustSlideExpressions(newSlidePart, context);

        // Clone slide layout relationship
        if (sourceSlidePart.SlideLayoutPart != null)
        {
            newSlidePart.AddPart(sourceSlidePart.SlideLayoutPart);
        }

        // Clone notes relationship and content if present
        if (sourceSlidePart.NotesSlidePart != null)
        {
            CloneNotesPartWithContext(sourceSlidePart, newSlidePart, context);
        }

        // Clone other important relationships (images, charts, etc.)
        CloneSlideRelationships(sourceSlidePart, newSlidePart);

        return newSlidePart;
    }

    /// <summary>
    /// Adjusts expression indices in all text elements of a slide
    /// </summary>
    public static void AdjustSlideExpressions(SlidePart slidePart, Models.SlideContext context)
    {
        if (slidePart == null || context == null || context.Offset == 0)
            return;

        Logger.Debug($"Adjusting expressions in slide for context: {context.GetContextDescription()}");

        try
        {
            // Find all text elements in the slide
            var textElements = slidePart.Slide.Descendants<D.Text>().ToList();
            Logger.Debug($"Found {textElements.Count} text elements to adjust");

            foreach (var textElement in textElements)
            {
                string originalText = textElement.Text;
                if (string.IsNullOrEmpty(originalText))
                    continue;

                // Check if text contains expressions that need adjustment
                if (originalText.Contains("${") && originalText.Contains("["))
                {
                    string adjustedText = ExpressionAdjuster.AdjustExpressionIndices(originalText, context);

                    // Only update if adjusted text is different
                    if (adjustedText != originalText)
                    {
                        Logger.Debug($"Adjusted text expression: '{originalText}' -> '{adjustedText}'");
                        textElement.Text = adjustedText;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error adjusting slide expressions: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Clones the notes slide part
    /// </summary>
    private static void CloneNotesPart(SlidePart sourceSlidePart, SlidePart newSlidePart)
    {
        NotesSlidePart sourceNotesPart = sourceSlidePart.NotesSlidePart;
        NotesSlidePart newNotesPart = newSlidePart.AddNewPart<NotesSlidePart>();

        // Clone notes content
        newNotesPart.NotesSlide = (P.NotesSlide)sourceNotesPart.NotesSlide.CloneNode(true);
    }

    /// <summary>
    /// Clones the notes slide part and adjusts expressions
    /// </summary>
    private static void CloneNotesPartWithContext(SlidePart sourceSlidePart, SlidePart newSlidePart, Models.SlideContext context)
    {
        NotesSlidePart sourceNotesPart = sourceSlidePart.NotesSlidePart;
        NotesSlidePart newNotesPart = newSlidePart.AddNewPart<NotesSlidePart>();

        // Clone notes content
        newNotesPart.NotesSlide = (P.NotesSlide)sourceNotesPart.NotesSlide.CloneNode(true);

        // Adjust expressions in notes text elements
        if (context != null)
        {
            try
            {
                var noteTextElements = newNotesPart.NotesSlide.Descendants<D.Text>().ToList();
                foreach (var textElement in noteTextElements)
                {
                    string originalText = textElement.Text;
                    if (string.IsNullOrEmpty(originalText) || !originalText.Contains("${") || !originalText.Contains("["))
                        continue;

                    string adjustedText = ExpressionAdjuster.AdjustExpressionIndices(originalText, context);
                    if (adjustedText != originalText)
                    {
                        Logger.Debug($"Adjusted notes text expression: '{originalText}' -> '{adjustedText}'");
                        textElement.Text = adjustedText;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error adjusting notes expressions: {ex.Message}", ex);
            }
        }
    }

    /// <summary>
    /// Clones slide relationships (images, charts)
    /// </summary>
    private static void CloneSlideRelationships(SlidePart sourceSlidePart, SlidePart newSlidePart)
    {
        // Handle common relationships like images
        foreach (var idPartPair in sourceSlidePart.Parts)
        {
            OpenXmlPart part = idPartPair.OpenXmlPart;

            // Skip layout and notes parts (already handled)
            if (part is SlideLayoutPart || part is NotesSlidePart)
                continue;

            // Handle image parts
            if (part is ImagePart)
            {
                var imagePart = part as ImagePart;
                var newImagePart = newSlidePart.AddImagePart(imagePart.ContentType);

                using (Stream sourceStream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
                using (Stream targetStream = newImagePart.GetStream(FileMode.Create, FileAccess.Write))
                {
                    sourceStream.CopyTo(targetStream);
                }

                // Update references in the new slide
                string oldId = sourceSlidePart.GetIdOfPart(imagePart);
                string newId = newSlidePart.GetIdOfPart(newImagePart);
                UpdateImageReferences(newSlidePart.Slide, oldId, newId);
            }
        }
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