namespace DocuChef.Presentation.Exceptions
{
    public class DirectiveParsingException : Exception
    {
        public DirectiveParsingException(string message) : base(message) { }
        public DirectiveParsingException(string message, Exception innerException) : base(message, innerException) { }
    }

    public class BindingExpressionException : Exception
    {
        public BindingExpressionException(string message) : base(message) { }
        public BindingExpressionException(string message, Exception innerException) : base(message, innerException) { }
    }

    public class SlideGenerationException : Exception
    {
        public SlideGenerationException(string message) : base(message) { }
        public SlideGenerationException(string message, Exception innerException) : base(message, innerException) { }
    }

    public class DataBindingException : Exception
    {
        public DataBindingException(string message) : base(message) { }
        public DataBindingException(string message, Exception innerException) : base(message, innerException) { }
    }
    public class DirectiveException : Exception
    {
        public DirectiveException(string message) : base(message) { }
        public DirectiveException(string message, Exception innerException) : base(message, innerException) { }
    }

    /// <summary>
    /// Exception thrown when an element should be hidden due to array bounds or other conditions
    /// </summary>
    public class ElementHideException : Exception
    {
        public ElementHideException(string message) : base(message) { }
        public ElementHideException(string message, Exception innerException) : base(message, innerException) { }
    }
}