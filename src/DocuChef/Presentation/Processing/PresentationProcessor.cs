using DocuChef.Presentation.Core;
using DocuChef.Presentation.Models;
using DocumentFormat.OpenXml.Presentation;

namespace DocuChef.Presentation.Processing;

/// <summary>
/// Processes a presentation plan to generate the final presentation
/// </summary>
internal class PresentationProcessor
{
    /// <summary>
    /// Processes the presentation based on the generation plan
    /// </summary>
    public void ProcessPresentation(PresentationDocument presentationDoc, PresentationPlan plan, object data)
    {
        if (presentationDoc.PresentationPart == null)
            throw new InvalidOperationException("Invalid presentation document");

        PresentationPart presentationPart = presentationDoc.PresentationPart;

        // Get the slide list
        SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
        if (slideIdList == null)
        {
            Logger.Warning("No slides found in the presentation");
            return;
        }

        // Process slides according to plan and get new slide list
        List<SlideId> newSlideIds = ProcessPlan(presentationPart, plan);

        // Replace the slide list
        ReplaceSlideList(slideIdList, newSlideIds);

        // Process data binding for each slide
        ProcessDataBinding(presentationPart, plan, newSlideIds, data);

        Logger.Info($"Presentation generation complete. Total slides: {newSlideIds.Count}");
    }

    /// <summary>
    /// Processes data binding for each slide
    /// </summary>
    private void ProcessDataBinding(PresentationPart presentationPart, PresentationPlan plan, List<SlideId> slideIds, object data)
    {
        var slideIndex = 0;
        foreach (var slideId in slideIds)
        {
            // Skip if we've reached the end of the plan
            if (slideIndex >= plan.Slides.Count)
                break;

            var slidePart = presentationPart.GetPartById(slideId.RelationshipId!) as SlidePart;
            if (slidePart == null)
                continue;

            var context = plan.Slides[slideIndex++].Context;
            DataBinder.BindDataWithContext(slidePart, context, data);
        }
    }

    /// <summary>
    /// Processes the plan and returns the new slide IDs
    /// </summary>
    private List<SlideId> ProcessPlan(PresentationPart presentationPart, PresentationPlan plan)
    {
        // Get original slide list
        List<SlideId> originalSlides = new List<SlideId>(
            presentationPart.Presentation.SlideIdList.Elements<SlideId>());

        // Create a mapping from relationship ID to SlideId for quick lookup
        Dictionary<string, SlideId> relIdToSlideId = originalSlides
            .Where(s => s.RelationshipId != null)
            .ToDictionary(s => s.RelationshipId.Value);

        // Create a new slide list
        List<SlideId> newSlideIds = new List<SlideId>();

        // Calculate next available slide ID
        uint nextSlideId = originalSlides.Count > 0
            ? originalSlides.Max(s => s.Id.Value) + 1
            : 256; // Start with a reasonable ID if no slides exist

        // Track slides that have been processed to avoid duplicates
        HashSet<string> processedRelIds = new HashSet<string>();

        // Log plan summary
        var planSummary = plan.GetSummary();
        Logger.Info($"Processing plan with {planSummary.TotalIncludedSlides} slides");

        // Process each slide in the plan in the order they were added
        // This preserves the original presentation order
        int processedCount = 0;
        foreach (var plannedSlide in plan.IncludedSlides)
        {
            try
            {
                ProcessPlannedSlide(
                    plannedSlide,
                    presentationPart,
                    relIdToSlideId,
                    newSlideIds,
                    ref nextSlideId,
                    processedRelIds);

                processedCount++;
                if (processedCount % 10 == 0)
                {
                    Logger.Info($"Processed {processedCount} of {planSummary.TotalIncludedSlides} slides");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error processing planned slide: {ex.Message}", ex);
            }
        }

        Logger.Info($"Successfully processed {processedCount} slides");
        return newSlideIds;
    }

    /// <summary>
    /// Processes a single planned slide
    /// </summary>
    private void ProcessPlannedSlide(
        PlannedSlide plannedSlide,
        PresentationPart presentationPart,
        Dictionary<string, SlideId> relIdToSlideId,
        List<SlideId> newSlideIds,
        ref uint nextSlideId,
        HashSet<string> processedRelIds)
    {
        // Skip slides that are marked to be skipped
        if (plannedSlide.Operation == SlideOperation.Skip)
        {
            Logger.Debug($"Skipping slide with ID {plannedSlide.Source.SlideId}");
            return;
        }

        Logger.Debug($"Processing planned slide: {plannedSlide.Operation}, context: {plannedSlide.Context?.GetContextDescription() ?? "No context"}");

        // Find the original slide ID
        string relId = plannedSlide.Source.RelationshipId;
        if (!relIdToSlideId.TryGetValue(relId, out SlideId originalSlideId))
        {
            Logger.Warning($"Cannot find slide with relationship ID: {relId}");
            return;
        }

        // Get the slide part
        SlidePart originalSlidePart = (SlidePart)presentationPart.GetPartById(originalSlideId.RelationshipId.Value);

        // Determine if we need to clone or reuse the slide
        if (plannedSlide.Operation == SlideOperation.Clone)
        {
            CloneSlideWithExpressionAdjustment(
                plannedSlide,
                presentationPart,
                originalSlidePart,
                newSlideIds,
                ref nextSlideId);
        }
        else // Keep the original slide
        {
            KeepOriginalSlide(
                plannedSlide,
                originalSlidePart,
                originalSlideId,
                newSlideIds,
                processedRelIds);
        }
    }

    /// <summary>
    /// Clones a slide with expression adjustment and adds it to the presentation
    /// </summary>
    private void CloneSlideWithExpressionAdjustment(
        PlannedSlide plannedSlide,
        PresentationPart presentationPart,
        SlidePart originalSlidePart,
        List<SlideId> newSlideIds,
        ref uint nextSlideId)
    {
        Logger.Debug($"Cloning slide ID {plannedSlide.Source.SlideId} with expression adjustment");

        // Clone the slide with context for expression adjustment
        SlidePart newSlidePart = SlideManager.CloneSlideWithContext(
            presentationPart,
            originalSlidePart,
            plannedSlide.Context);

        string newRelId = presentationPart.GetIdOfPart(newSlidePart);

        // Create new slide ID
        SlideId newSlideId = new SlideId
        {
            Id = nextSlideId++,
            RelationshipId = newRelId
        };

        // Add to new slide list
        newSlideIds.Add(newSlideId);

        // Update slide notes with context
        SlideManager.UpdateSlideNote(newSlidePart, plannedSlide.Context);

        Logger.Debug($"Cloned slide ID {plannedSlide.Source.SlideId} -> {newSlideId.Id} with adjusted expressions");
    }

    /// <summary>
    /// Keeps the original slide and adds it to the presentation
    /// </summary>
    private void KeepOriginalSlide(
        PlannedSlide plannedSlide,
        SlidePart originalSlidePart,
        SlideId originalSlideId,
        List<SlideId> newSlideIds,
        HashSet<string> processedRelIds)
    {
        string relId = originalSlideId.RelationshipId.Value;

        // Only add original slide if not already processed
        if (!processedRelIds.Contains(relId))
        {
            newSlideIds.Add(originalSlideId);
            processedRelIds.Add(relId);

            // Update slide notes with context
            SlideManager.UpdateSlideNote(originalSlidePart, plannedSlide.Context);

            // Adjust expressions in the original slide based on context
            if (plannedSlide.HasContext)
            {
                // For original slides we still want to adjust expressions
                SlideManager.AdjustSlideExpressions(originalSlidePart, plannedSlide.Context);
            }

            Logger.Debug($"Kept original slide ID {originalSlideId.Id} with adjusted expressions");
        }
        else
        {
            Logger.Debug($"Skipping duplicate original slide ID {originalSlideId.Id}");
        }
    }

    /// <summary>
    /// Replaces the slide list in the presentation
    /// </summary>
    private void ReplaceSlideList(SlideIdList slideIdList, List<SlideId> newSlideIds)
    {
        // Remove all existing slides
        slideIdList.RemoveAllChildren();

        // Add new slides in order
        foreach (SlideId id in newSlideIds)
        {
            slideIdList.Append(id);
        }
    }
}