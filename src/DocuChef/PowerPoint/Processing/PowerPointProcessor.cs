using DocuChef.PowerPoint.DollarSignEngine;
using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing
{
    /// <summary>
    /// Main coordinator for PowerPoint template processing
    /// </summary>
    internal class PowerPointProcessor : IExpressionEvaluator
    {
        private readonly PresentationDocument _document;
        private readonly PowerPointOptions _options;
        private readonly PowerPointContext _context;
        private readonly ExpressionEvaluator _expressionEvaluator;

        // Specialized processors
        private readonly SlideProcessor _slideProcessor;
        private readonly BindingProcessor _bindingProcessor;
        private readonly DirectiveProcessor _directiveProcessor;
        private readonly ExpressionProcessor _expressionProcessor;

        /// <summary>
        /// Initialize PowerPoint processor
        /// </summary>
        public PowerPointProcessor(PresentationDocument document, PowerPointOptions options)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _options = options ?? throw new ArgumentNullException(nameof(options));

            _context = new PowerPointContext { Options = options };
            _expressionEvaluator = new ExpressionEvaluator();

            // Initialize sub-processors
            _slideProcessor = new SlideProcessor(this, _context);
            _bindingProcessor = new BindingProcessor(this, _context);
            _directiveProcessor = new DirectiveProcessor(this, _context);
            _expressionProcessor = new ExpressionProcessor(this, _context);

            Logger.Debug("PowerPoint processor initialized");
        }

        /// <summary>
        /// Process PowerPoint template with variables and functions - simplified flow
        /// </summary>
        public void Process(Dictionary<string, object> variables,
                    Dictionary<string, Func<object>> globalVariables,
                    Dictionary<string, PowerPointFunction> functions)
        {
            // 1. Initialization
            InitializeContext(variables, globalVariables, functions);
            var presentationPart = ValidateDocument();
            var slideIds = GetSlideIds(presentationPart);

            Logger.Info("Phase 1: Analyzing and preparing slides...");

            // 2. Analyze and duplicate slides (slide preparation phase)
            foreach (var slideId in slideIds.ToList())  // Copy the collection as it may change during iteration
            {
                var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                _slideProcessor.AnalyzeAndPrepareSlide(presentationPart, slidePart);
            }

            // 3. Apply bindings to all prepared slides
            Logger.Info("Phase 2: Applying bindings to all slides...");
            var allSlideIds = GetSlideIds(presentationPart);
            foreach (var slideId in allSlideIds)
            {
                var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                _bindingProcessor.ApplyBindings(slidePart);
            }

            // Final save
            presentationPart.Presentation.Save();
            Logger.Info("PowerPoint template processing completed successfully");
        }

        /// <summary>
        /// Initialize context with variables and functions
        /// </summary>
        private void InitializeContext(Dictionary<string, object> variables, Dictionary<string, Func<object>> globalVariables, Dictionary<string, PowerPointFunction> functions)
        {
            _context.Variables = variables ?? new Dictionary<string, object>();
            _context.GlobalVariables = globalVariables ?? new Dictionary<string, Func<object>>();
            _context.Functions = functions ?? new Dictionary<string, PowerPointFunction>();
            _context.Variables["_context"] = _context;
        }

        /// <summary>
        /// Validate document structure
        /// </summary>
        private PresentationPart ValidateDocument()
        {
            var presentationPart = _document.PresentationPart;
            if (presentationPart?.Presentation?.SlideIdList == null)
            {
                throw new DocuChefException("Invalid PowerPoint document structure");
            }

            return presentationPart;
        }

        /// <summary>
        /// Get slide IDs from presentation
        /// </summary>
        private List<SlideId> GetSlideIds(PresentationPart presentationPart)
        {
            return presentationPart.Presentation.SlideIdList
                .ChildElements.OfType<SlideId>()
                .ToList();
        }

        /// <summary>
        /// Prepare variables dictionary
        /// </summary>
        internal Dictionary<string, object> PrepareVariables()
        {
            var variables = new Dictionary<string, object>(_context.Variables);

            // Add global variables
            foreach (var globalVar in _context.GlobalVariables)
            {
                variables[globalVar.Key] = globalVar.Value();
            }

            // Add PowerPoint functions
            foreach (var function in _context.Functions)
            {
                variables[$"ppt.{function.Key}"] = function.Value;
            }

            return variables;
        }

        /// <summary>
        /// Evaluate expression with provided variables
        /// </summary>
        public object EvaluateCompleteExpression(string expression, Dictionary<string, object> variables)
        {
            try
            {
                // If already wrapped in ${...}, evaluate directly
                if (expression.StartsWith("${") && expression.EndsWith("}"))
                {
                    var result = _expressionEvaluator.Evaluate(expression, variables);
                    return result;
                }

                // Otherwise, wrap it for evaluation
                string wrappedExpr = "${" + expression + "}";
                var evalResult = _expressionEvaluator.Evaluate(wrappedExpr, variables);
                return evalResult;
            }
            catch (Exception ex)
            {
                Logger.Error($"Error evaluating expression '{expression}': {ex.Message}", ex);
                return $"[Error: {ex.Message}]";
            }
        }
    }
}