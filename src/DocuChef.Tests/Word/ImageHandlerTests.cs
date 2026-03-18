using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocuChef.Word.Models;
using DocuChef.Word.Processors;
using FluentAssertions;
using Xunit;
using WDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocuChef.Tests.Word;

public class ImageHandlerTests : IDisposable
{
    private readonly string _tempPngPath;

    public ImageHandlerTests()
    {
        // Minimal valid 1x1 PNG
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==");
        _tempPngPath = Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid():N}.png");
        File.WriteAllBytes(_tempPngPath, png);
    }

    public void Dispose()
    {
        if (File.Exists(_tempPngPath))
            File.Delete(_tempPngPath);
    }

    [Fact]
    public void PlaceholderInRun_InsertsDrawingElement()
    {
        // Arrange
        const string placeholderKey = "%%IMG_logo%%";
        using var stream = WordTestHelper.CreateDocx(placeholderKey);
        using var doc = WordprocessingDocument.Open(stream, true);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document.Body!;

        var images = new Dictionary<string, ImagePlaceholder>
        {
            [placeholderKey] = new ImagePlaceholder { Path = _tempPngPath }
        };

        // Act
        ImageHandler.ProcessImages(mainPart, body, images);

        // Assert
        var drawings = body.Descendants<WDrawing>().ToList();
        drawings.Should().HaveCount(1, "the placeholder run should be replaced with a Drawing element");

        // The placeholder text should no longer exist
        var remainingTexts = body.Descendants<WText>()
            .Select(t => t.Text)
            .Where(t => t.Contains(placeholderKey));
        remainingTexts.Should().BeEmpty("placeholder text should be removed after image insertion");
    }

    [Fact]
    public void NoPlaceholders_NoChange()
    {
        // Arrange
        using var stream = WordTestHelper.CreateDocx("Normal text without placeholders");
        using var doc = WordprocessingDocument.Open(stream, true);
        var mainPart = doc.MainDocumentPart!;
        var body = mainPart.Document.Body!;

        var images = new Dictionary<string, ImagePlaceholder>
        {
            ["%%IMG_logo%%"] = new ImagePlaceholder { Path = _tempPngPath }
        };

        // Act
        ImageHandler.ProcessImages(mainPart, body, images);

        // Assert
        var drawings = body.Descendants<WDrawing>().ToList();
        drawings.Should().BeEmpty("no placeholder text exists, so no Drawing should be inserted");

        var text = body.Descendants<WText>().First().Text;
        text.Should().Be("Normal text without placeholders");
    }
}
