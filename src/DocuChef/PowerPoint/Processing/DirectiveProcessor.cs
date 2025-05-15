using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing
{
    /// <summary>
    /// Processor responsible for handling directives from slide notes
    /// </summary>
    internal class DirectiveProcessor
    {
        private readonly PowerPointProcessor _mainProcessor;
        private readonly PowerPointContext _context;

        /// <summary>
        /// Initialize directive processor
        /// </summary>
        public DirectiveProcessor(PowerPointProcessor mainProcessor, PowerPointContext context)
        {
            _mainProcessor = mainProcessor ?? throw new ArgumentNullException(nameof(mainProcessor));
            _context = context ?? throw new ArgumentNullException(nameof(context));
        }

        /// <summary>
        /// Process shape directive
        /// </summary>
        public void ProcessShapeDirective(SlidePart slidePart, DirectiveContext directive)
        {
            _context.Directive = directive;

            // Get target shape name
            if (!directive.Parameters.TryGetValue("target", out var targetName) || string.IsNullOrEmpty(targetName))
            {
                Logger.Warning($"No target shape specified for directive {directive.Name}");
                return;
            }

            // Clean target name
            targetName = CleanDirectiveParameter(targetName);

            // Find target shapes
            var targetShapes = FindShapesByName(slidePart, targetName);
            if (!targetShapes.Any())
            {
                Logger.Warning($"Target shape '{targetName}' not found");
                return;
            }

            // Process directive based on type
            switch (directive.Name)
            {
                case "if":
                    ProcessIfDirective(targetShapes, directive);
                    break;
                default:
                    Logger.Warning($"Unknown directive: {directive.Name}");
                    break;
            }
        }

        /// <summary>
        /// Process if directive
        /// </summary>
        private void ProcessIfDirective(List<P.Shape> targetShapes, DirectiveContext directive)
        {
            string condition = directive.Value.Trim();
            Logger.Debug($"Processing if directive with condition: {condition}");

            try
            {
                // Evaluate condition
                var result = EvaluateDirectiveCondition(condition);
                bool conditionResult = ConvertToBoolean(result);

                Logger.Debug($"Condition evaluated to: {conditionResult}");

                // Set visibility of target shapes using PowerPointShapeHelper
                foreach (var shape in targetShapes)
                {
                    if (conditionResult)
                        PowerPointShapeHelper.ShowShape(shape);
                    else
                        PowerPointShapeHelper.HideShape(shape);
                }

                // Handle visibleWhenFalse shapes if specified
                if (directive.Parameters.TryGetValue("visibleWhenFalse", out var visibleWhenFalseName))
                {
                    visibleWhenFalseName = CleanDirectiveParameter(visibleWhenFalseName);
                    var visibleWhenFalseShapes = FindShapesByName(_context.SlidePart, visibleWhenFalseName);

                    // Set opposite visibility for these shapes
                    foreach (var shape in visibleWhenFalseShapes)
                    {
                        if (!conditionResult)
                            PowerPointShapeHelper.ShowShape(shape);
                        else
                            PowerPointShapeHelper.HideShape(shape);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error processing if directive: {condition}", ex);
            }
        }

        /// <summary>
        /// Evaluate directive condition
        /// </summary>
        private object EvaluateDirectiveCondition(string expression)
        {
            if (string.IsNullOrEmpty(expression))
                return false;

            try
            {
                var variables = _mainProcessor.PrepareVariables();
                return _mainProcessor.EvaluateCompleteExpression(expression, variables);
            }
            catch (Exception ex)
            {
                Logger.Error($"Error evaluating directive condition '{expression}': {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// Find shapes by name
        /// </summary>
        private List<P.Shape> FindShapesByName(SlidePart slidePart, string targetName)
        {
            var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();
            var targetShapes = new List<P.Shape>();

            foreach (var shape in shapes)
            {
                string shapeName = shape.GetShapeName();

                if (shapeName == targetName)
                {
                    targetShapes.Add(shape);
                    Logger.Debug($"Found shape '{targetName}'");
                }
            }

            return targetShapes;
        }

        /// <summary>
        /// Convert result to boolean
        /// </summary>
        private bool ConvertToBoolean(object result)
        {
            if (result is bool boolValue)
                return boolValue;

            if (result != null)
            {
                try
                {
                    return Convert.ToBoolean(result);
                }
                catch
                {
                    return false;
                }
            }

            return false;
        }

        /// <summary>
        /// Clean directive parameter by removing quotes
        /// </summary>
        private string CleanDirectiveParameter(string parameter)
        {
            parameter = parameter.Trim();

            if (parameter.StartsWith("\"") && parameter.EndsWith("\"") && parameter.Length > 1)
            {
                parameter = parameter.Substring(1, parameter.Length - 2);
            }

            return parameter;
        }
    }
}