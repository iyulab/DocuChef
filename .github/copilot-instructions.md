# DocuChef Project Instructions

This workspace contains the DocuChef library for document processing with PowerPoint and Excel support.

## Development Guidelines

Please follow the guidelines in these files when working on this project:

1. **AI Development Rules**: See `src/DocuChef/AI_RULE.md`
   - Design-first approach for PPT
   - Fundamental problem-solving over patches
   - Minimal design patterns, avoid over-engineering
   - Code deduplication through Helpers, Extensions, inheritance
   - Refactor when files exceed 500 lines
   - Concise code, meaningful naming, minimal comments
   - Conservative access modifiers (internal), careful use of public
   - No hardcoding for generic library solutions

2. **PowerPoint Design Architecture**: See `src/DocuChef/PPT_DESIGN.md`
   - 4-stage processing flow: Analysis → Planning → Generation → Binding
   - Core components: TemplateAnalyzer, SlidePlanGenerator, SlideGenerator, DataBinder
   - Key models: SlideInfo, SlidePlan, SlideInstance, Directive, BindingExpression

3. **PowerPoint Template Syntax**: See `src/DocuChef/SYNTAX_OF_PPT.md`
   - Design-first approach
   - Value binding with `${...}` syntax
   - Contextual hierarchy with `>` operator
   - Control directives in slide notes

## Code Standards

- Write code and comments in English
- Provide responses in Korean when requested
- Show complete code without omissions
- Use code artifacts when appropriate
