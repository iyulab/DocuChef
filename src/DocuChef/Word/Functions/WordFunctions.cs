using DocuChef.Word.Models;

namespace DocuChef.Word.Functions;

public class WordFunctions
{
    private readonly string? _templateDirectory;

    public WordFunctions(string? templateDirectory = null)
    {
        _templateDirectory = templateDirectory;
    }

    public ImagePlaceholder Image(string path, int? widthPx = null, int? heightPx = null)
    {
        string resolvedPath = path;
        if (!string.IsNullOrEmpty(_templateDirectory) &&
            !Path.IsPathRooted(path) &&
            !path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
            !path.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            resolvedPath = Path.Combine(_templateDirectory, path);
        }

        const long emuPerPixel = 9525;
        return new ImagePlaceholder
        {
            Path = resolvedPath,
            Width = widthPx.HasValue ? widthPx.Value * emuPerPixel : null,
            Height = heightPx.HasValue ? heightPx.Value * emuPerPixel : null,
        };
    }
}
