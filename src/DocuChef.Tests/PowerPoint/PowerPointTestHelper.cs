using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace DocuChef.Tests.PowerPoint;

/// <summary>
/// Creates and reads .pptx documents programmatically for testing.
/// </summary>
public static class PowerPointTestHelper
{
    /// <summary>
    /// Creates a minimal pptx with one slide containing the given text in a text box.
    /// </summary>
    public static MemoryStream CreatePptx(string slideText)
        => CreatePptx(new[] { slideText });

    /// <summary>
    /// Creates a minimal pptx with one slide per text entry.
    /// </summary>
    public static MemoryStream CreatePptx(IEnumerable<string> slideTexts)
    {
        var stream = new MemoryStream();
        using (var prs = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presPart = prs.AddPresentationPart();
            presPart.Presentation = new P.Presentation();

            var masterPart = presPart.AddNewPart<SlideMasterPart>();
            masterPart.SlideMaster = new SlideMaster(
                new P.CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()))));

            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            layoutPart.SlideLayout = new SlideLayout(
                new P.CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()))));
            layoutPart.SlideLayout.Save();
            masterPart.SlideMaster.Save();

            var slideIdList = new SlideIdList();
            presPart.Presentation.Append(slideIdList);
            presPart.Presentation.SlideSize = new SlideSize { Cx = 9144000, Cy = 5143500 };
            presPart.Presentation.NotesSize = new NotesSize { Cx = 6858000, Cy = 9144000 };

            uint slideId = 256;
            foreach (var text in slideTexts)
            {
                var slidePart = presPart.AddNewPart<SlidePart>();
                slidePart.AddPart(layoutPart);
                slidePart.Slide = BuildSlide(text);
                slidePart.Slide.Save();

                string relId = presPart.GetIdOfPart(slidePart);
                slideIdList.Append(new SlideId { Id = slideId++, RelationshipId = relId });
            }

            presPart.Presentation.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Collects all text from all slides.
    /// </summary>
    public static List<string> ReadAllText(Stream stream)
    {
        stream.Position = 0;
        using var prs = PresentationDocument.Open(stream, false);
        var texts = new List<string>();

        var slideIdList = prs.PresentationPart?.Presentation?.SlideIdList;
        if (slideIdList == null) return texts;

        foreach (SlideId sid in slideIdList.Elements<SlideId>())
        {
            var relId = sid.RelationshipId?.Value;
            if (relId == null) continue;
            var slidePart = (SlidePart)prs.PresentationPart!.GetPartById(relId);
            foreach (var t in slidePart.Slide!.Descendants<D.Text>())
                texts.Add(t.Text);
        }
        return texts;
    }

    /// <summary>
    /// Creates a pptx with a single slide carrying one text box per entry, so that
    /// each expression lives in its own shape (the design-first layout).
    /// </summary>
    public static MemoryStream CreatePptxWithShapes(params string[] shapeTexts)
    {
        var stream = new MemoryStream();
        using (var prs = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presPart = prs.AddPresentationPart();
            presPart.Presentation = new P.Presentation();

            var masterPart = presPart.AddNewPart<SlideMasterPart>();
            masterPart.SlideMaster = new SlideMaster(
                new P.CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()))));

            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            layoutPart.SlideLayout = new SlideLayout(
                new P.CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()))));
            layoutPart.SlideLayout.Save();
            masterPart.SlideMaster.Save();

            var slideIdList = new SlideIdList();
            presPart.Presentation.Append(slideIdList);
            presPart.Presentation.SlideSize = new SlideSize { Cx = 9144000, Cy = 5143500 };
            presPart.Presentation.NotesSize = new NotesSize { Cx = 6858000, Cy = 9144000 };

            var slidePart = presPart.AddNewPart<SlidePart>();
            slidePart.AddPart(layoutPart);

            var tree = new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new TransformGroup()));

            uint shapeId = 2;
            foreach (var text in shapeTexts)
                tree.Append(BuildShape(text, shapeId++));

            slidePart.Slide = new Slide(new P.CommonSlideData(tree));
            slidePart.Slide.Save();

            slideIdList.Append(new SlideId { Id = 256, RelationshipId = presPart.GetIdOfPart(slidePart) });
            presPart.Presentation.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Creates a pptx whose slides carry control directives in their notes. Directives are
    /// only read from notes, never from slide content.
    /// </summary>
    public static MemoryStream CreatePptxWithNotes(params (string SlideText, string Notes)[] slides)
    {
        var stream = new MemoryStream();
        using (var prs = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presPart = prs.AddPresentationPart();
            presPart.Presentation = new P.Presentation();

            var masterPart = presPart.AddNewPart<SlideMasterPart>();
            masterPart.SlideMaster = new SlideMaster(
                new P.CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()))));

            var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            layoutPart.SlideLayout = new SlideLayout(
                new P.CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()))));
            layoutPart.SlideLayout.Save();
            masterPart.SlideMaster.Save();

            var slideIdList = new SlideIdList();
            presPart.Presentation.Append(slideIdList);
            presPart.Presentation.SlideSize = new SlideSize { Cx = 9144000, Cy = 5143500 };
            presPart.Presentation.NotesSize = new NotesSize { Cx = 6858000, Cy = 9144000 };

            uint slideId = 256;
            foreach (var (slideText, notes) in slides)
            {
                var slidePart = presPart.AddNewPart<SlidePart>();
                slidePart.AddPart(layoutPart);
                slidePart.Slide = BuildSlide(slideText);
                slidePart.Slide.Save();

                if (!string.IsNullOrEmpty(notes))
                {
                    var notesPart = slidePart.AddNewPart<NotesSlidePart>();
                    var notesTree = new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new TransformGroup()));

                    // One paragraph per line: the analyzer reads directives line by line.
                    var body = new P.TextBody(new BodyProperties(), new ListStyle());
                    foreach (var line in notes.Split('\n'))
                        body.Append(new D.Paragraph(new D.Run(new D.Text(line.TrimEnd('\r')))));

                    notesTree.Append(new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 2, Name = "Notes Placeholder" },
                            new P.NonVisualShapeDrawingProperties(new ShapeLocks { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape { Type = PlaceholderValues.Body })),
                        new P.ShapeProperties(),
                        body));

                    notesPart.NotesSlide = new NotesSlide(
                        new P.CommonSlideData(notesTree),
                        new ColorMapOverride(new D.MasterColorMapping()));
                    notesPart.NotesSlide.Save();
                }

                slideIdList.Append(new SlideId { Id = slideId++, RelationshipId = presPart.GetIdOfPart(slidePart) });
            }

            presPart.Presentation.Save();
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Counts the image parts embedded across all slides — the evidence that a picture
    /// was actually inserted rather than merely referenced.
    /// </summary>
    public static int CountEmbeddedImages(Stream stream)
    {
        stream.Position = 0;
        using var prs = PresentationDocument.Open(stream, false);

        var slideIdList = prs.PresentationPart?.Presentation?.SlideIdList;
        if (slideIdList == null) return 0;

        int count = 0;
        foreach (SlideId sid in slideIdList.Elements<SlideId>())
        {
            var relId = sid.RelationshipId?.Value;
            if (relId == null) continue;
            var slidePart = (SlidePart)prs.PresentationPart!.GetPartById(relId);
            count += slidePart.ImageParts.Count();
        }
        return count;
    }

    /// <summary>
    /// Writes a minimal valid 1x1 PNG to <paramref name="path"/>.
    /// </summary>
    public static void WriteTinyPng(string path) =>
        File.WriteAllBytes(path, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));

    /// <summary>
    /// Collects text per slide, preserving slide order. Index 0 is the first slide.
    /// </summary>
    public static List<string> ReadTextBySlide(Stream stream)
    {
        stream.Position = 0;
        using var prs = PresentationDocument.Open(stream, false);
        var perSlide = new List<string>();

        var slideIdList = prs.PresentationPart?.Presentation?.SlideIdList;
        if (slideIdList == null) return perSlide;

        foreach (SlideId sid in slideIdList.Elements<SlideId>())
        {
            var relId = sid.RelationshipId?.Value;
            if (relId == null) continue;
            var slidePart = (SlidePart)prs.PresentationPart!.GetPartById(relId);
            perSlide.Add(string.Concat(slidePart.Slide!.Descendants<D.Text>().Select(t => t.Text)));
        }
        return perSlide;
    }

    private static P.Shape BuildShape(string text, uint shapeId) =>
        new(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox{shapeId}" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
            new P.ShapeProperties(),
            new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new D.Paragraph(new D.Run(new D.Text(text)))));

    private static Slide BuildSlide(string text)
    {
        var sp = BuildShape(text, 2);

        return new Slide(
            new P.CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()),
                    sp)));
    }
}
