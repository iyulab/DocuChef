namespace DocuChef.Presentation.Exceptions
{
    public class DirectiveParsingException : DocuChefException
    {
        public DirectiveParsingException(string message) : base(message) { }
        public DirectiveParsingException(string message, Exception innerException) : base(message, innerException) { }
    }

    public class BindingExpressionException : DocuChefException
    {
        public BindingExpressionException(string message) : base(message) { }
        public BindingExpressionException(string message, Exception innerException) : base(message, innerException) { }
    }

    public class SlideGenerationException : DocuChefException
    {
        public SlideGenerationException(string message) : base(message) { }
        public SlideGenerationException(string message, Exception innerException) : base(message, innerException) { }
    }

    public class DataBindingException : DocuChefException
    {
        public DataBindingException(string message) : base(message) { }
        public DataBindingException(string message, Exception innerException) : base(message, innerException) { }
    }

    public class DirectiveException : DocuChefException
    {
        public DirectiveException(string message) : base(message) { }
        public DirectiveException(string message, Exception innerException) : base(message, innerException) { }
    }

    /// <summary>
    /// Exception thrown when an element should be hidden due to array bounds or other conditions.
    /// Alias for DocuChefHideException — use DocuChefHideException for new code.
    /// </summary>
    public class ElementHideException : DocuChefHideException
    {
        public ElementHideException(string message) : base(message) { }
        public ElementHideException(string message, Exception innerException) : base(message, innerException) { }
    }
}
