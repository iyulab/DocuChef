
namespace DocuChef.Presentation;

public class PowerPointRecipe : IRecipe
{
    private string templatePath;
    private PowerPointOptions options;

    public PowerPointRecipe(string templatePath, PowerPointOptions options)
    {
        this.templatePath = templatePath;
        this.options = options;
    }

    public PowerPointRecipe(Stream templateStream, PowerPointOptions powerPointOptions)
    {
    }

    public void AddVariable(string name, object value)
    {
        throw new NotImplementedException();
    }

    public void AddVariable(object data)
    {
        throw new NotImplementedException();
    }

    public void ClearVariables()
    {
        throw new NotImplementedException();
    }

    public void Dispose()
    {
        throw new NotImplementedException();
    }

    public void RegisterGlobalVariable(string name, object value)
    {
        throw new NotImplementedException();
    }

    internal IDish Generate()
    {
        throw new NotImplementedException();
    }
}
