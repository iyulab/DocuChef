# PowerPoint Template Syntax Guide

## Overview

The PowerPoint Template Syntax provides a simple and intuitive way to bind data to your designed PowerPoint slides. This template engine preserves the original design of your presentation while dynamically populating it with your data.

## Design Principles

1. **Design-First Approach**: PowerPoint is design-centered, so the templating system binds data to pre-designed elements
2. **Template Preservation**: Maintains the original PowerPoint design and formatting
3. **Automatic Processing**: Intelligent detection of patterns reduces manual directive requirements
4. **Error Resilience**: Graceful handling of missing data and invalid expressions
5. **Performance Optimized**: Efficient processing with caching and smart algorithms

## Syntax Structure

### 1. Value Binding (Inside Slide Elements)

#### Basic Syntax
```
${PropertyName}                     // Basic property binding
${Object.PropertyName}              // Nested property binding
${Value:FormatSpecifier}            // Using format specifiers
${Array[Index].PropertyName}        // Array element property binding
```

#### Advanced Syntax
```
${Condition ? Value1 : Value2}      // Conditional expressions (if supported)
${Method()}                         // Method calls (if supported)
${Parent.Child.PropertyName}        // Deep nested property binding
${Parent.Child[0].PropertyName}     // Combined nested and array access
```

### 2. Contextual Hierarchy with '>' Operator

The `>` operator enables contextual data access, essential for multi-slide generation with nested collections:

```
${Parent>Child.PropertyName}              // Contextual relationship using '>' operator
${Categories>Items[0].Name}               // First item name in the current category
${Departments>Teams[0]>Members[2].Name}   // Third member name in the first team of current department
```

**Key Differences**:
- **Dot Notation (.)**: Always references from the absolute path starting from root data
  - Example: `Categories[0].Items[0].Name` always refers to first item of first category
- **Context Operator (>)**: References from the current item being processed
  - Example: `Categories>Items[0].Name` refers to first item of current category in each slide

### 3. Special Functions (Inside Slide Elements)

#### Image Binding
```
${ppt.Image("ImageProperty")}                    // Basic image binding
${ppt.Image("Product.Photo")}                    // Image from nested property
${ppt.Image("Product.Photo", width: 300, height: 200, preserveAspectRatio: true)}  // With options
```

#### Function Parameters
- **ImageProperty**: Path to image data (URL, file path, or base64)
- **width**: Desired width in pixels (optional)
- **height**: Desired height in pixels (optional)
- **preserveAspectRatio**: Maintain original aspect ratio (optional, default: true)

### 4. Control Directives (In Slide Notes Only)

Control directives are placed in slide notes and provide explicit control over slide generation:

#### #foreach Directive
```
#foreach: Collection                               // Basic iteration
#foreach: Collection, max: Number                  // With max items per slide
#foreach: Collection, max: Number, offset: Number  // With max items and offset
```

#### #range Directive
```
#range-begin: Collection    // Start of a grouped range
#range-end: Collection      // End of a grouped range
```

#### #alias Directive
```
#alias: FullPath as ShortName    // Create path alias for simpler expressions
```

**Important Notes**:
- Directives are **optional** - the engine automatically detects patterns
- Manual directives override automatic detection
- All directives must be placed in slide notes, not in slide content

## Automatic Array Processing

The engine automatically detects and processes array patterns without requiring explicit directives:

### 1. Pattern Detection
The engine analyzes `${Array[Index].Property}` patterns in slides:
- **Index Analysis**: Finds highest array index in expressions
- **Items Per Slide**: Calculated as `MaxIndex + 1`
- **Required Slides**: `TotalItems ÷ ItemsPerSlide`

### 2. Automatic Slide Generation
```
// Example: If slide contains:
Product 1: ${Products[0].Name} - ${Products[0].Price}
Product 2: ${Products[1].Name} - ${Products[1].Price}
Product 3: ${Products[2].Name} - ${Products[2].Price}

// Result: 3 items per slide (indices 0, 1, 2)
// If Products array has 8 items:
//   Slide 1: Items 0, 1, 2
//   Slide 2: Items 3, 4, 5  
//   Slide 3: Items 6, 7, (empty)
```

### 3. Automatic Index Adjustment
- **First Slide**: `Products[0]`, `Products[1]`, `Products[2]`
- **Second Slide**: Expressions automatically become `Products[3]`, `Products[4]`, `Products[5]`
- **Third Slide**: Expressions become `Products[6]`, `Products[7]`, empty for missing data

## Nested Collections and Context

### Contextual Processing
For complex data structures with parent-child relationships:

```json
{
  "Categories": [
    {
      "Name": "Electronics",
      "Items": [
        { "Name": "Smartphone", "Price": 999 },
        { "Name": "Laptop", "Price": 1299 },
        { "Name": "Tablet", "Price": 599 }
      ]
    },
    {
      "Name": "Furniture", 
      "Items": [
        { "Name": "Sofa", "Price": 799 },
        { "Name": "Table", "Price": 499 }
      ]
    }
  ]
}
```

### Slide Design Example
```
Title: ${Categories[0].Name} Category
Item 1: ${Categories>Items[0].Name}: $${Categories>Items[0].Price}
Item 2: ${Categories>Items[1].Name}: $${Categories>Items[1].Price}
Item 3: ${Categories>Items[2].Name}: $${Categories>Items[2].Price}
```

### Generated Results
- **Slide 1** (Categories[0] = Electronics):
  - Title: "Electronics Category"
  - Item 1: "Smartphone: $999"
  - Item 2: "Laptop: $1299"
  - Item 3: "Tablet: $599"

- **Slide 2** (Categories[1] = Furniture):
  - Title: "Furniture Category"
  - Item 1: "Sofa: $799"
  - Item 2: "Table: $499"
  - Item 3: (empty)

## Format Specifiers

Format specifiers control how values are displayed:

### Numeric Formatting
```
${Price:C}           // Currency: $1,299.00
${Price:N2}          // Number with 2 decimals: 1,299.00
${Price:F0}          // Fixed point, no decimals: 1299
${Percentage:P}      // Percentage: 85.50%
```

### Date Formatting
```
${Date:yyyy-MM-dd}   // ISO format: 2024-03-15
${Date:MMMM dd, yyyy} // Long format: March 15, 2024
${Date:M/d/yyyy}     // Short format: 3/15/2024
```

### String Formatting
```
${Name:U}            // Uppercase
${Name:L}            // Lowercase  
${Description:20}    // Truncate to 20 characters
```

## Empty Value Handling

When expressions cannot be resolved (missing data, out-of-bounds indices):

1. **Text Elements**: Replaced with empty string ("")
2. **Image Elements**: 
   - Use placeholder image if available
   - Hide element if no placeholder
   - Log warning for debugging
3. **Conditional Display**: Elements can be hidden based on data availability

## Advanced Features

### Alias Usage
Create shorter aliases for complex paths:

```
// In slide notes:
#alias: Company.Departments.Engineering.Employees as Engineers

// In slide content:
Manager: ${Engineers[0].Name}
Team Lead: ${Engineers[1].Name}
```

### Range Processing
Group related slides for batch processing:

```
// Slide 1 notes:
#range-begin: Products

// Slides 2-5 content:
Product: ${Products>Details[0].Name}
Price: ${Products>Details[0].Price}

// Slide 6 notes:
#range-end: Products
```

### Conditional Processing
Handle optional data gracefully:

```
${Product.Name ?? "No Product Name"}           // Default value
${Product.Image || "default-image.png"}        // Alternative value
${Customer.Phone ? Customer.Phone : "N/A"}     // Ternary operator
```

## Best Practices

### Template Design
1. **Start with Design**: Create the PowerPoint layout first, then add expressions
2. **Consistent Patterns**: Use consistent array indexing across related slides
3. **Clear Structure**: Organize data logically to match slide flow
4. **Test with Sample Data**: Validate templates with representative data

### Expression Writing
1. **Simple Paths**: Keep data paths as straightforward as possible
2. **Meaningful Names**: Use descriptive property names in data objects
3. **Handle Missing Data**: Design for cases where data might be incomplete
4. **Format Appropriately**: Apply formatting at the expression level

### Performance Optimization
1. **Limit Expressions**: Avoid excessive binding expressions per slide
2. **Optimize Data Structure**: Organize data to match template needs
3. **Use Caching**: Leverage built-in caching for repeated data access
4. **Monitor Size**: Be mindful of large datasets and memory usage

## Syntax Placement Rules

### Value Bindings & Special Functions
- **Location**: Inside text content of slide elements (text boxes, shapes, tables, charts)
- **Method**: Select element in PowerPoint → Enter text editing mode → Add expressions
- **Scope**: Can be mixed with regular text content

### Control Directives  
- **Location**: Only in slide notes (View → Notes in PowerPoint)
- **Format**: Each directive on a separate line
- **Scope**: Applies to the entire slide

### Element Naming
- **Purpose**: Reference specific shapes in directives
- **Method**: Select shape → Right-click → "Edit Alt Text" or use Selection Pane
- **Usage**: Reference by name in control directives when needed

## Error Handling and Debugging

### Common Issues
1. **Expression Not Found**: Check data structure matches expression path
2. **Array Index Errors**: Verify indices don't exceed available data
3. **Context Resolution**: Ensure parent-child relationships are correctly structured
4. **Missing Images**: Verify image paths and file accessibility

### Debug Information
The engine provides detailed logging for troubleshooting:
- Expression parsing results
- Data resolution steps  
- Slide generation progress
- Binding success/failure details

### Validation Tools
- **Expression Tester**: Test expressions against sample data
- **Template Analyzer**: Review detected patterns and directives
- **Data Inspector**: Examine data structure and types
- **Performance Monitor**: Track processing time and memory usage

## Migration Guide

### From DollarSignEngine
The syntax is largely compatible with existing DollarSignEngine templates:
- Most `${...}` expressions work unchanged
- New contextual `>` operator provides enhanced capabilities
- Automatic directive generation reduces manual configuration
- Improved error handling and performance

### Upgrading Templates
1. **Review Expressions**: Check for any syntax changes needed
2. **Test Directives**: Verify explicit directives still work as expected
3. **Add Context**: Consider using `>` operator for nested collections
4. **Optimize Performance**: Remove unnecessary directives (rely on auto-detection)

## Examples

### Simple Product List
```
// Slide content:
Title: Product Catalog
Item 1: ${Products[0].Name} - $${Products[0].Price:N2}
Item 2: ${Products[1].Name} - $${Products[1].Price:N2}
Item 3: ${Products[2].Name} - $${Products[2].Price:N2}

// Result: Automatically creates multiple slides for all products, 3 per slide
```

### Employee Directory with Departments
```
// Slide content:
Department: ${Departments[0].Name}
Manager: ${Departments>Employees[0].Name}
Staff: ${Departments>Employees[1].Name}, ${Departments>Employees[2].Name}

// Result: One slide per department with department-specific employees
```

### Sales Report with Images
```
// Slide content:
Quarter: ${Reports[0].Quarter}
Revenue: ${Reports[0].Revenue:C}
Chart: ${ppt.Image("Reports[0].ChartImage")}
Growth: ${Reports[0].Growth:P1}

// Result: One slide per quarter with embedded charts
```
