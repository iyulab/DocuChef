using System;

namespace DocuChef.Exceptions;

/// <summary>
/// Base exception for all DocuChef errors
/// </summary>
public class DocuChefException : Exception
{
    /// <summary>
    /// Creates a new DocuChef exception with the specified message
    /// </summary>
    public DocuChefException(string message) : base(message)
    {
    }

    /// <summary>
    /// Creates a new DocuChef exception with the specified message and inner exception
    /// </summary>
    public DocuChefException(string message, Exception innerException) : base(message, innerException)
    {
    }
}

/// <summary>
/// Exception thrown when a template format is invalid or unsupported
/// </summary>
public class InvalidTemplateFormatException : DocuChefException
{
    /// <summary>
    /// Creates a new invalid template format exception with the specified message
    /// </summary>
    public InvalidTemplateFormatException(string message) : base(message)
    {
    }

    /// <summary>
    /// Creates a new invalid template format exception with the specified message and inner exception
    /// </summary>
    public InvalidTemplateFormatException(string message, Exception innerException) : base(message, innerException)
    {
    }
}

/// <summary>
/// Exception thrown when a template processing error occurs
/// </summary>
public class TemplateProcessingException : DocuChefException
{
    /// <summary>
    /// Creates a new template processing exception with the specified message
    /// </summary>
    public TemplateProcessingException(string message) : base(message)
    {
    }

    /// <summary>
    /// Creates a new template processing exception with the specified message and inner exception
    /// </summary>
    public TemplateProcessingException(string message, Exception innerException) : base(message, innerException)
    {
    }
}

/// <summary>
/// Exception thrown when a variable operation fails
/// </summary>
public class VariableOperationException : DocuChefException
{
    /// <summary>
    /// Creates a new variable operation exception with the specified message
    /// </summary>
    public VariableOperationException(string message) : base(message)
    {
    }

    /// <summary>
    /// Creates a new variable operation exception with the specified message and inner exception
    /// </summary>
    public VariableOperationException(string message, Exception innerException) : base(message, innerException)
    {
    }
}

/// <summary>
/// Exception thrown when a PowerPoint element should be hidden due to out-of-range array access
/// </summary>
public class DocuChefHideException : DocuChefException
{
    /// <summary>
    /// Creates a new hide exception for PowerPoint elements
    /// </summary>
    public DocuChefHideException() : base("Element should be hidden due to out-of-range array access")
    {
    }

    /// <summary>
    /// Creates a new hide exception with the specified message
    /// </summary>
    public DocuChefHideException(string message) : base(message)
    {
    }

    /// <summary>
    /// Creates a new hide exception with the specified message and inner exception
    /// </summary>
    public DocuChefHideException(string message, Exception innerException) : base(message, innerException)
    {
    }
}