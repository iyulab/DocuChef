using DocuChef.Word.Models;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;

namespace DocuChef.Word.Processors;

/// <summary>
/// Replaces image placeholder markers with actual OpenXML Drawing elements in Word documents.
/// </summary>
public static class ImageHandler
{
    // Default image size: 4 inches x 3 inches in EMUs (1 inch = 914400 EMU)
    private const long DefaultWidthEmu = 914400L * 4;
    private const long DefaultHeightEmu = 914400L * 3;

    private static int _imageCounter;

    /// <summary>
    /// Finds runs containing placeholder keys and replaces them with Drawing elements
    /// containing the referenced images.
    /// </summary>
    public static void ProcessImages(
        MainDocumentPart mainPart,
        OpenXmlElement container,
        Dictionary<string, ImagePlaceholder> images)
    {
        if (images.Count == 0)
            return;

        var runs = container.Descendants<W.Run>().ToList();

        foreach (var run in runs)
        {
            var textElement = run.GetFirstChild<W.Text>();
            if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                continue;

            var text = textElement.Text;

            foreach (var (key, placeholder) in images)
            {
                if (text != key)
                    continue;

                if (!File.Exists(placeholder.Path))
                {
                    Logger.Debug($"ImageHandler: Image file not found: {placeholder.Path}");
                    continue;
                }

                // Add the image part
                var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (var fileStream = File.OpenRead(placeholder.Path))
                {
                    imagePart.FeedData(fileStream);
                }

                var relationshipId = mainPart.GetIdOfPart(imagePart);
                var width = placeholder.Width ?? DefaultWidthEmu;
                var height = placeholder.Height ?? DefaultHeightEmu;

                // Build the Drawing element and replace the run's content
                var drawing = CreateDrawingElement(relationshipId, width, height);

                // Remove text content from the run and insert Drawing
                textElement.Remove();
                run.Append(drawing);

                break; // This run matched; move to next run
            }
        }
    }

    private static WDrawing CreateDrawingElement(string relationshipId, long widthEmu, long heightEmu)
    {
        var imageId = Interlocked.Increment(ref _imageCounter);
        var elementId = (uint)imageId;

        var drawing = new WDrawing(
            new DW.Inline(
                new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                new DW.DocProperties { Id = elementId, Name = $"Image{imageId}" },
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = widthEmu, Cy = heightEmu }
                                ),
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = A.ShapeTypeValues.Rectangle
                                }
                            )
                        )
                    )
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            }
        );

        return drawing;
    }
}
