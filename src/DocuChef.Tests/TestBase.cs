using Xunit.Abstractions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;

namespace DocuChef.Tests;

/// <summary>
/// Base class for XLCustom test classes with common helper methods
/// </summary>
public abstract class TestBase : IDisposable
{
    protected readonly Xunit.Abstractions.ITestOutputHelper _output;

    protected TestBase(Xunit.Abstractions.ITestOutputHelper output)
    {
        _output = output ?? throw new ArgumentNullException(nameof(output));
    }

    public virtual void Dispose()
    {
        GC.SuppressFinalize(this);
    }

    public Chef CreateNewChef()
    {
        var chef = new Chef(new RecipeOptions()
        {
            EnableVerboseLogging = true,
        });
        return chef;
    }

    /// <summary>
    /// Creates a mock slide part for testing purposes
    /// </summary>
    /// <param name="slideContent">The text content to place in the slide</param>
    /// <returns>A mock slide part object</returns>
    protected MockSlidePart CreateMockSlidePart(string slideContent)
    {
        return new MockSlidePart(slideContent);
    }
}

/// <summary>
/// Mock slide part for testing purposes
/// </summary>
public class MockSlidePart
{
    public string Content { get; }

    public MockSlidePart(string content)
    {
        Content = content ?? string.Empty;
    }
}