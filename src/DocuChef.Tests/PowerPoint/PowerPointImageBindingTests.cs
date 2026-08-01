using FluentAssertions;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

/// <summary>
/// Black-box coverage for <c>${ppt.Image(...)}</c>. The expression resolves to an internal
/// marker during binding; a second pass swaps the marker's shape for a real picture part.
/// Contract source: SYNTAX_OF_PPT.md "Image Binding" and "Empty Value Handling", which
/// specifies a different overflow rule for images than for text — hide, not blank.
/// </summary>
public class PowerPointImageBindingTests : TestBase
{
    public PowerPointImageBindingTests(ITestOutputHelper output) : base(output) { }

    private (List<string> Texts, int Images) Render(string slideText, string variableName, object variableValue)
    {
        string tempPath = TempPptx();
        try
        {
            using var stream = PowerPointTestHelper.CreatePptx(slideText);
            File.WriteAllBytes(tempPath, stream.ToArray());

            using var chef = CreateNewChef();
            using var recipe = chef.LoadTemplate(tempPath);
            recipe.AddVariable(variableName, variableValue);
            using var dish = recipe.Generate();

            var result = new MemoryStream();
            dish.SaveAs(result);

            var texts = PowerPointTestHelper.ReadTextBySlide(result);
            int images = PowerPointTestHelper.CountEmbeddedImages(result);
            _output.WriteLine($"texts = [{string.Join(" | ", texts)}], embedded images = {images}");
            return (texts, images);
        }
        finally { Cleanup(tempPath); }
    }

    [Fact]
    public void ImageBinding_ExistingFile_EmbedsPicture()
    {
        string imagePath = TempPng();
        try
        {
            PowerPointTestHelper.WriteTinyPng(imagePath);

            var (texts, images) = Render("${ppt.Image(Logo)}", "Logo", imagePath);

            images.Should().Be(1, "the resolved image must be embedded as a picture part");
            texts.Should().NotContain(t => t.Contains("__PPT_IMAGE_"),
                "the internal marker must never survive into the generated document");
            texts.Should().NotContain(t => t.Contains("${"),
                "the raw expression must never survive into the generated document");
        }
        finally { Cleanup(imagePath); }
    }

    /// <summary>
    /// Underscores are ordinary in file names ("company_logo.png") and appear in Windows temp
    /// paths. The internal marker is delimited by "__", so a path carried inside it used to
    /// make the marker unparseable — the image silently vanished and the marker shipped.
    /// </summary>
    [Fact]
    public void ImageBinding_PathContainsUnderscore_StillEmbedsPicture()
    {
        string imagePath = Path.Combine(Path.GetTempPath(), $"docuchef_logo_{Guid.NewGuid():N}.png");
        try
        {
            PowerPointTestHelper.WriteTinyPng(imagePath);

            var (texts, images) = Render("${ppt.Image(Logo)}", "Logo", imagePath);

            images.Should().Be(1, "an underscore in the path must not defeat marker resolution");
            texts.Should().NotContain(t => t.Contains("__PPT_IMAGE_"));
        }
        finally { Cleanup(imagePath); }
    }

    [Fact]
    public void ImageBinding_MissingFile_LeavesNoMarkerBehind()
    {
        string missingPath = TempPng(); // deliberately never created

        var (texts, images) = Render("${ppt.Image(Logo)}", "Logo", missingPath);

        images.Should().Be(0, "a missing source cannot produce a picture part");
        texts.Should().NotContain(t => t.Contains("__PPT_IMAGE_"),
            "an unresolvable image must not leave its internal marker in the document");
        texts.Should().NotContain(t => t.Contains("${"),
            "an unresolvable image must not leave the raw expression in the document");
    }

    private static string TempPptx() =>
        Path.Combine(Path.GetTempPath(), $"docuchef_test_{Guid.NewGuid():N}.pptx");

    // No underscore in the file name: the marker payload is delimited by "__", so an
    // underscore in the path is itself a variable under test (see UnderscoreInPath test).
    private static string TempPng() =>
        Path.Combine(Path.GetTempPath(), $"docuchef{Guid.NewGuid():N}.png");

    private static void Cleanup(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }
}
