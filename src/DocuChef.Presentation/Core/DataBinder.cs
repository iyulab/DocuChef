using DocuChef.Presentation.Handlers;
using DollarSignEngine;

namespace DocuChef.Presentation.Core;

/// <summary>
/// Handles data binding to slide elements using DollarSignEngine
/// </summary>
internal class DataBinder
{
    /// <summary>
    /// Binds data to all text elements in a slide using slide context
    /// </summary>
    public static void BindDataWithContext(SlidePart slidePart, Models.SlideContext context, object data)
    {
        if (slidePart == null || context == null)
        {
            Logger.Debug("Cannot bind data: slidePart or context is null");
            return;
        }

        try
        {
            Logger.Debug($"Starting data binding for slide with context: {context.GetContextDescription()}");

            // Create DollarSignOptions with context-aware variable resolver
            var options = DollarSignOptions.Default
                .WithDollarSignSyntax()
                .WithGlobalData(data)
                .WithErrorHandler((expr, ex) =>
                {
                    Logger.Debug($"Expression error: '{expr}' - {ex.Message}");
                    return string.Empty;  // Return empty on error
                });

            // Bind all text elements
            BindTextElements(slidePart, context, options);

            Logger.Debug("Data binding completed successfully");
        }
        catch (Exception ex)
        {
            Logger.Error($"Error binding data to slide: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Binds text elements in a slide
    /// </summary>
    private static void BindTextElements(SlidePart slidePart, SlideContext context, DollarSignOptions options)
    {
        var textElements = slidePart.Slide.Descendants<D.Text>().ToList();
        Logger.Debug($"Found {textElements.Count} text elements in slide");

        foreach (var textElement in textElements)
        {
            string originalText = textElement.Text;
            if (string.IsNullOrEmpty(originalText))
                continue;

            // Check if text contains potential bindings
            if (originalText.Contains('{') || originalText.Contains('$'))
            {
                try
                {
                    var data = context.GetData();
                    if (data is IEnumerable enumerable
                        && int.TryParse(originalText.GetBetween("[", "]"), out var ndx)
                        && context.TotalItems <= ndx)
                    {
                        textElement.Hide();
                        continue;
                    }
                    // Use DollarSignEngine to evaluate the text
                    string newText = DollarSign.Eval(originalText,
                        data,
                        options);

                    // Only update if the text actually changed
                    if (newText != originalText)
                    {
                        if (originalText.Contains($"ppt.{nameof(PPTMethods.Image)}"))
                        {
                            var shape = textElement.FindShape();
                            if (shape != null)
                            {
                                ImageHandler.Process(shape, newText);
                            }
                        }

                        Logger.Debug($"Binding text: '{originalText}' -> '{newText}'");
                        textElement.Text = newText;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error($"Error evaluating expression '{originalText}': {ex.Message}");
                }
            }
        }
    }
}