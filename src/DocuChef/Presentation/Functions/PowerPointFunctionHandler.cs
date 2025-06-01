using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using DocuChef.Logging;
using DocuChef.Presentation.Handlers;

namespace DocuChef.Presentation.Functions;

/// <summary>
/// Handles PowerPoint-specific function processing like image insertion
/// </summary>
public static class PowerPointFunctionHandler
{
    private static readonly Regex ImagePlaceholderRegex = new(@"__PPT_IMAGE_([a-f0-9]{32})__([^_]+)(?:__(\d+)__(\d+)__(True|False))?__", RegexOptions.Compiled);

    /// <summary>
    /// <param name="presentationDocument">The presentation document</param>
    /// <param name="pptFunctions">The PPT functions instance with cached data</param>
    public static void ProcessFunctions(PresentationDocument presentationDocument, PPTFunctions pptFunctions)
    {
        Logger.Debug("PowerPointFunctionHandler.ProcessFunctions started");

        if (presentationDocument?.PresentationPart?.Presentation?.SlideIdList == null)
        {
            Logger.Debug("PowerPointFunctionHandler: No presentation or slide list found");
            return;
        }

        var slideIds = presentationDocument.PresentationPart.Presentation.SlideIdList.ChildElements
            .OfType<SlideId>().ToList();

        Logger.Debug($"PowerPointFunctionHandler: Found {slideIds.Count} slides to process");

        foreach (var slideId in slideIds)
        {
            if (string.IsNullOrEmpty(slideId.RelationshipId?.Value))
            {
                Logger.Debug("PowerPointFunctionHandler: Skipping slide with null RelationshipId");
                continue;
            }

            try
            {
                Logger.Debug($"PowerPointFunctionHandler: Processing slide with RelationshipId: {slideId.RelationshipId.Value}");
                var slidePart = (SlidePart)presentationDocument.PresentationPart.GetPartById(slideId.RelationshipId.Value);
                ProcessSlideImagePlaceholders(slidePart, pptFunctions);
            }
            catch (Exception ex)
            {
                Logger.Error($"Error processing slide functions: {ex.Message}", ex);
            }
        }

        Logger.Debug("PowerPointFunctionHandler.ProcessFunctions completed");
    }

    /// <summary>
    /// Processes image placeholders in a single slide
    /// </summary>
    /// <param name="slidePart">The slide part to process</param>
    /// <param name="pptFunctions">The PPT functions instance with cached data</param>
    private static void ProcessSlideImagePlaceholders(SlidePart slidePart, PPTFunctions pptFunctions)
    {
        Logger.Debug("ProcessSlideImagePlaceholders started");

        if (slidePart?.Slide == null)
        {
            Logger.Debug("ProcessSlideImagePlaceholders: slidePart or Slide is null");
            return;
        }

        var imageCache = pptFunctions.GetAllImageCache();
        Logger.Debug($"ProcessSlideImagePlaceholders: Retrieved image cache with {imageCache.Count} items");

        if (imageCache.Count == 0)
        {
            Logger.Debug("ProcessSlideImagePlaceholders: No cached images found, exiting");
            return;
        }

        Logger.Debug($"Processing image placeholders in slide. Found {imageCache.Count} cached images.");

        // Find all text elements that might contain image placeholders
        var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
        Logger.Debug($"ProcessSlideImagePlaceholders: Found {textElements.Count} text elements");

        foreach (var textElement in textElements)
        {
            if (string.IsNullOrEmpty(textElement.Text))
                continue;

            Logger.Debug($"ProcessSlideImagePlaceholders: Checking text element: '{textElement.Text}'");

            var matches = ImagePlaceholderRegex.Matches(textElement.Text);
            Logger.Debug($"ProcessSlideImagePlaceholders: Found {matches.Count} placeholder matches");

            foreach (Match match in matches)
            {
                Logger.Debug($"ProcessSlideImagePlaceholders: Processing match: {match.Value}");
                ProcessImagePlaceholder(slidePart, textElement, match, imageCache);
            }
        }

        Logger.Debug("ProcessSlideImagePlaceholders completed");
    }

    /// <summary>
    /// Processes a single image placeholder
    /// </summary>
    /// <param name="slidePart">The slide part</param>
    /// <param name="textElement">The text element containing the placeholder</param>
    /// <param name="match">The regex match for the placeholder</param>
    /// <param name="imageCache">The image cache dictionary</param>
    private static void ProcessImagePlaceholder(SlidePart slidePart, DocumentFormat.OpenXml.Drawing.Text textElement, Match match, Dictionary<string, string> imageCache)
    {
        try
        {
            var placeholder = match.Value;
            var guid = match.Groups[1].Value;
            var propertyPath = match.Groups[2].Value;

            Logger.Debug($"Processing image placeholder: {placeholder}");

            // Check if we have cached image data for this placeholder
            if (!imageCache.TryGetValue(placeholder, out var imageFilePath))
            {
                Logger.Warning($"No cached image data found for placeholder: {placeholder}");
                return;
            }

            // Parse dimensions if provided
            int width = 300; // Default width
            int height = 200; // Default height
            bool preserveAspectRatio = true;

            if (match.Groups.Count > 3 && !string.IsNullOrEmpty(match.Groups[3].Value))
            {
                int.TryParse(match.Groups[3].Value, out width);
                int.TryParse(match.Groups[4].Value, out height);
                bool.TryParse(match.Groups[5].Value, out preserveAspectRatio);
            }

            // Find the shape that contains this text element
            var shape = FindContainingShape(textElement);
            if (shape == null)
            {
                Logger.Warning($"Could not find containing shape for image placeholder: {placeholder}");
                return;
            }

            Logger.Debug($"Found containing shape for image: {shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value}");

            // Process the image using ImageHandler with the actual file path
            ImageHandler.Process(shape, imageFilePath);

            // Clear the text element that contained the placeholder
            textElement.Text = string.Empty;

            Logger.Debug($"Successfully processed image placeholder: {placeholder}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing image placeholder: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Finds the shape that contains the given text element
    /// </summary>
    /// <param name="textElement">The text element</param>
    /// <returns>The containing shape or null if not found</returns>
    private static DocumentFormat.OpenXml.Presentation.Shape? FindContainingShape(DocumentFormat.OpenXml.Drawing.Text textElement)
    {
        var current = textElement.Parent;
        while (current != null)
        {
            if (current is DocumentFormat.OpenXml.Presentation.Shape shape)
                return shape;
            current = current.Parent;
        }
        return null;
    }

    /// <summary>
    /// Cleans up any temporary resources used by the function handler
    /// </summary>
    public static void Cleanup()
    {
        ImageHandler.Cleanup();
    }
}