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

        var slideIdList = prs.PresentationPart?.Presentation.SlideIdList;
        if (slideIdList == null) return texts;

        foreach (SlideId sid in slideIdList.Elements<SlideId>())
        {
            var relId = sid.RelationshipId?.Value;
            if (relId == null) continue;
            var slidePart = (SlidePart)prs.PresentationPart!.GetPartById(relId);
            foreach (var t in slidePart.Slide.Descendants<D.Text>())
                texts.Add(t.Text);
        }
        return texts;
    }

    private static Slide BuildSlide(string text)
    {
        var txBody = new P.TextBody(
            new BodyProperties(),
            new ListStyle(),
            new D.Paragraph(new D.Run(new D.Text(text))));

        var sp = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2, Name = "TextBox" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
            new P.ShapeProperties(),
            txBody);

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
