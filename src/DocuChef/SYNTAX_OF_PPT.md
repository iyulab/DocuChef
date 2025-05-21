# PowerPoint Template Syntax Guide

## Overview

The PowerPoint Template Syntax provides a simple and intuitive way to bind data to your designed PowerPoint slides. This template engine preserves the original design of your presentation while dynamically populating it with your data.

## Basic Principles

1. **Design-First Approach**: PowerPoint is design-centered, so the templating system binds data to pre-designed elements.
2. **DollarSignEngine Compatible**: Maintains familiarity for users of the existing DollarSignEngine.
3. **Simplicity**: Intuitive and streamlined syntax for ease of use.
4. **Design Preservation**: Respects the original intent of PowerPoint design elements.

## Syntax Structure

### 1. Value Binding (Inside Slide Elements)
```
${PropertyName}                // Basic property binding
${Object.PropertyName}         // Nested property binding
${Value:FormatSpecifier}       // Using format specifiers
${Condition ? Value1 : Value2} // Conditional expressions
${Method()}                    // Method calls
${Array[Index].PropertyName}   // Array element property binding
${Parent.Child.PropertyName}   // Parent-child relationship property binding
${Parent.Child[0].PropertyName} // Parent-child with array element binding
```

### 2. Contextual Hierarchy with '>' Operator
```
${Parent>Child.PropertyName}              // Contextual relationship using '>' operator
${Categories>Item[0].Name}                // First item name in the current category
${Departments>Teams[0]>Members[2].Name}   // Third member name in the first team of the current department
```

The '>' operator is used to reference items in the current processing context, which is especially useful for multi-slide generation:

- **Dot Notation (.)**: Always references from the absolute path starting from the root of the data.
  Example: `Categories[0].Items[0].Name` always refers to the first item of the first category.

- **Context Operator (>)**: References from the current item being processed.
  Example: `Categories>Items[0].Name` refers to the first item of the current category in each slide.

### 3. Special Functions (Inside Slide Elements)
```
${ppt.Image("ImageProperty")}   // Image binding
${ppt.Image("Product.Photo", width: 300, height: 200, preserveAspectRatio: true)}
```

### 4. Control Directives (In Slide Notes Only)
```
#foreach: Collection, max: Number, offset: Number // Array iteration (optional)
#range-begin: Collection // Start of a group area for an array
#range-end: Collection // End of a group area for an array
```

**Note**: These directives are typically handled automatically by the engine through design analysis. Manual entry is optional and will override automatic detection if provided.

## Array Data Processing

PowerPoint's design-centered approach means array processing binds data to pre-designed elements:

### Automatic Array Indexing

1. **Direct Array Element Reference**: Use array indices directly in slide elements
   ```
   Product 1: ${Products[0].Name} - ${Products[0].Price}
   Product 2: ${Products[1].Name} - ${Products[1].Price}
   Product 3: ${Products[2].Name} - ${Products[2].Price}
   ```

2. **Automatic Slide Generation**: The engine automatically duplicates slides based on array size
   - Example: If the `Products` array has 8 items:
     - First slide: Items 0, 1, 2
     - Second slide: Items 3, 4, 5
     - Third slide: Items 6, 7

3. **Automatic Index Offset Calculation**: Indices are automatically adjusted in duplicated slides
   - First slide: `Products[0]`, `Products[1]`, `Products[2]`
   - Second slide: Same references automatically convert to `Products[3]`, `Products[4]`, `Products[5]`
   - Third slide: Same references automatically convert to `Products[6]`, `Products[7]`, empty value

## How It Works

1. **Array Index Pattern Detection**: The engine automatically analyzes `${Array[Index].Property}` patterns in the slide
2. **Items Per Slide Calculation**: Determined by the highest index + 1 in a slide
   - Example: If a slide has `[0]`, `[1]`, `[2]`, then 3 items per slide
3. **Required Slides Calculation**: `Total Items ÷ Items Per Slide`
4. **Automatic Slide Duplication**: Original slide is duplicated as needed
5. **Automatic Index Adjustment**: Index references are automatically adjusted in each duplicated slide

## The #foreach Directive (Optional)

The `#foreach` directive provides an explicit method for array data processing but is **not required**. The library can automatically analyze design elements to detect and process array patterns.

### #foreach Syntax
```
#foreach: Collection, max: Number, offset: Number
```

- **Collection**: Array or collection to iterate (required)
- **max**: Maximum items per slide (optional, default: auto-detect)
- **offset**: Starting index offset (optional, default: 0)

### Automatic Detection vs. Explicit #foreach

1. **Automatic Detection (Default Behavior)**
   - The library automatically detects array patterns even without the `#foreach` directive
   - Analyzes `${Array[Index]}` patterns in the slide and automatically duplicates slides as needed
   - Automatically adjusts indices in duplicated slides

2. **Explicit #foreach (Optional)**
   - Use the `#foreach` directive for more granular control
   - Explicitly specify items per slide
   - Specify starting offset
   - Useful for nested array processing

## Examples

### Basic Presentation Slide

**Slide Element Content:**
- Title text box: `${Report.Title}`
- Subtitle text box: `${Report.Subtitle}`
- Date text box: `${Report.Date:yyyy-MM-dd}`
- Logo image: `${ppt.Image("Company.Logo")}`

### Product List Slide (Using Array Indices)

**Slide Element Content:**
- Title text box: `${Category.Name} Products`
- Product Item 1: 
  ```
  ${Products[0].Id}. ${Products[0].Name}
  Price: $${Products[0].Price:N2}
  ```
- Product Item 2: 
  ```
  ${Products[1].Id}. ${Products[1].Name}
  Price: $${Products[1].Price:N2}
  ```
- Product Item 3: 
  ```
  ${Products[2].Id}. ${Products[2].Name}
  Price: $${Products[2].Price:N2}
  ```

**Slide Notes (Optional):**
```
#foreach: Products, max: 3  # Optional: The library automatically detects the index pattern
```

**Result:**
- If the `Products` array has 8 items:
  - First slide: Items 0, 1, 2
  - Second slide: Items 3, 4, 5
  - Third slide: Items 6, 7

### Department Team Members Slide (Using Contextual '>' Operator)

**Slide Element Content:**
- Title text box: `${Departments[0].Name} Department`
- Manager text box: `Manager: ${Departments[0].Manager}`
- Headcount text box: `Headcount: ${Departments[0].Members.Length}`
- Team Member 1: `${Departments>Members[0].Name} (${Departments>Members[0].Position})`
- Team Member 2: `${Departments>Members[1].Name} (${Departments>Members[1].Position})`

**Result:**
- If there are 2 departments with 5 and 3 members respectively:
  - First slide: Department 1, Members 0, 1
  - Second slide: Department 1, Members 2, 3
  - Third slide: Department 1, Member 4 (one empty slot)
  - Fourth slide: Department 2, Members 0, 1
  - Fifth slide: Department 2, Member 2 (one empty slot)

## Nested Context Example

For multi-level nested data like:

```json
{
  "Categories": [
    {
      "Name": "Electronics",
      "Items": [
        { "Name": "Smartphone", "Price": 999 },
        { "Name": "Laptop", "Price": 1299 }
      ]
    },
    {
      "Name": "Furniture",
      "Items": [
        { "Name": "Sofa", "Price": 799 },
        { "Name": "Dining Table", "Price": 499 }
      ]
    }
  ]
}
```

**Slide Design:**
- Title: `${Categories[0].Name} Category`
- Item List:
  - `${Categories>Items[0].Name}: $${Categories>Items[0].Price}`
  - `${Categories>Items[1].Name}: $${Categories>Items[1].Price}`

**Result:**
- First slide (Categories[0]):
  - Title: "Electronics Category"
  - Item List:
    - Smartphone: $999
    - Laptop: $1299

- Second slide (Categories[1]):
  - Title: "Furniture Category"
  - Item List:
    - Sofa: $799
    - Dining Table: $499

## Empty Value Handling

When array indices exceed the actual data range:
1. **Text Elements**: Replaced with empty string ("")
2. **Image Elements**: Replaced with default image or hidden
3. **Chart/Table Elements**: Displayed as no-data state or hidden

## Syntax Placement

1. **Value Bindings & Special Functions**: 
   - Place inside text content of slide elements like text boxes, shapes, tables, charts
   - Select the element in PowerPoint and enter in text editing mode

2. **Control Directives**: 
   - Place only in slide notes (View > Notes)
   - Place each directive on a new line if multiple are needed

3. **Element Identification**: 
   - In PowerPoint, select a shape > right-click > Name Shape
   - Reference by assigned name in control directives