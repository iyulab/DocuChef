namespace DocuChef.PowerPoint.Processing.Directives;

/// <summary>
/// Interface for directive handlers
/// </summary>
internal interface IDirectiveHandler
{
    /// <summary>
    /// Checks if this handler can process the given directive
    /// </summary>
    bool CanHandle(SlideDirective directive);

    /// <summary>
    /// Processes the directive
    /// </summary>
    SlideProcessingResult Process(
        PresentationPart presentationPart,
        SlidePart slidePart,
        SlideDirective directive,
        PowerPointContext context,
        Dictionary<string, object> variables);
}