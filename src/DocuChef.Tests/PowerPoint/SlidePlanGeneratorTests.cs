using System.Collections.Generic;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Processors;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint;

/// <summary>
/// Tests for SlidePlanGenerator - Step 2: Slide Plan Generation
/// Validates slide instance planning based on template analysis and data
/// </summary>
public class SlidePlanGeneratorTests : TestBase
{
    public SlidePlanGeneratorTests(ITestOutputHelper output) : base(output) { }

    [Fact]
    public void GeneratePlan_SimpleArrayBinding_CalculatesCorrectSlideCount()
    {
        // Arrange
        var generator = new SlidePlanGenerator();
        var slideInfos = new List<SlideInfo>
        {
            new SlideInfo
            {
                SlideId = 0,
                Type = SlideType.Source,
                Position = 0,
                CollectionName = "Items",
                MaxArrayIndex = 1, // Items[0], Items[1] = 2 items per slide
                BindingExpressions = new List<BindingExpression>
                {
                    new BindingExpression { OriginalExpression = "${Items[0].Name}", DataPath = "Items[0].Name" },
                    new BindingExpression { OriginalExpression = "${Items[1].Name}", DataPath = "Items[1].Name" }
                }
            }
        };

        var data = new Dictionary<string, object>
        {
            ["Items"] = new[]
            {
                new { Name = "Item 1" },
                new { Name = "Item 2" },
                new { Name = "Item 3" },
                new { Name = "Item 4" },
                new { Name = "Item 5" }
            }
        };

        // Act
        var slidePlan = generator.GeneratePlan(slideInfos, data);

        // Assert
        slidePlan.Should().NotBeNull();
        slidePlan.SlideInstances.Should().HaveCount(3); // ⌈5 ÷ 2⌉ = 3 slides

        var instances = slidePlan.SlideInstances;
        instances[0].SourceSlideId.Should().Be(0);
        instances[0].IndexOffset.Should().Be(0);
        instances[0].Position.Should().Be(0);

        instances[1].SourceSlideId.Should().Be(0);
        instances[1].IndexOffset.Should().Be(2);
        instances[1].Position.Should().Be(1);

        instances[2].SourceSlideId.Should().Be(0);
        instances[2].IndexOffset.Should().Be(4);
        instances[2].Position.Should().Be(2);
    }

    [Fact]
    public void GeneratePlan_NestedCollections_HandlesContextOperator()
    {
        // Arrange
        var generator = new SlidePlanGenerator();
        var slideInfos = new List<SlideInfo>
        {
            // Category slide
            new SlideInfo
            {
                SlideId = 0,
                Type = SlideType.Source,
                Position = 0,
                CollectionName = "Categories",
                MaxArrayIndex = 0, // Categories[0] = 1 category per slide
                Directives = new List<Directive>
                {
                    new Directive { Type = DirectiveType.Range, RangeBoundary = RangeBoundary.Begin, SourceName = "Categories" }
                }
            },
            // Product slide
            new SlideInfo
            {
                SlideId = 1,
                Type = SlideType.Source,
                Position = 1,
                CollectionName = "Categories>Products",
                MaxArrayIndex = 1, // Products[0], Products[1] = 2 products per slide
                Directives = new List<Directive>
                {
                    new Directive { Type = DirectiveType.Range, RangeBoundary = RangeBoundary.End, SourceName = "Categories" }
                }
            }
        };

        var data = new Dictionary<string, object>
        {
            ["Categories"] = new[]
            {
                new 
                { 
                    Name = "Electronics", 
                    Products = new[]
                    {
                        new { Name = "Phone", Price = 1000 },
                        new { Name = "Tablet", Price = 800 },
                        new { Name = "Laptop", Price = 1500 }
                    }
                },
                new 
                { 
                    Name = "Furniture", 
                    Products = new[]
                    {
                        new { Name = "Sofa", Price = 500 },
                        new { Name = "Bed", Price = 700 }
                    }
                }
            }
        };

        // Act
        var slidePlan = generator.GeneratePlan(slideInfos, data);

        // Assert
        slidePlan.Should().NotBeNull();
        // Expected: 2 category slides + 2 product slides for Electronics + 1 product slide for Furniture = 5 total
        slidePlan.SlideInstances.Should().HaveCount(5);

        var categorySlides = slidePlan.SlideInstances.Where(s => s.SourceSlideId == 0).ToList();
        categorySlides.Should().HaveCount(2);

        var productSlides = slidePlan.SlideInstances.Where(s => s.SourceSlideId == 1).ToList();
        productSlides.Should().HaveCount(3); // 2 for Electronics + 1 for Furniture
    }

    [Fact]
    public void GeneratePlan_MixedStaticAndDynamic_PreservesStaticSlidePositions()
    {
        // Arrange
        var generator = new SlidePlanGenerator();
        var slideInfos = new List<SlideInfo>
        {
            new SlideInfo { SlideId = 0, Type = SlideType.Static, Position = 0 }, // Title
            new SlideInfo { SlideId = 1, Type = SlideType.Static, Position = 1 }, // Agenda
            new SlideInfo 
            { 
                SlideId = 2, 
                Type = SlideType.Source, 
                Position = 2, 
                CollectionName = "Items",
                MaxArrayIndex = 1 // 2 items per slide
            },
            new SlideInfo { SlideId = 3, Type = SlideType.Static, Position = 3 }, // Summary
        };

        var data = new Dictionary<string, object>
        {
            ["Items"] = new[] { new { Name = "Item 1" }, new { Name = "Item 2" }, new { Name = "Item 3" } }
        };

        // Act
        var slidePlan = generator.GeneratePlan(slideInfos, data);

        // Assert
        slidePlan.Should().NotBeNull();
        // Expected: Title(1) + Agenda(1) + Items(2) + Summary(1) = 5 slides
        slidePlan.SlideInstances.Should().HaveCount(5);

        var staticSlides = slidePlan.SlideInstances.Where(s => s.Type == SlideInstanceType.Static).ToList();
        staticSlides.Should().HaveCount(3); // Title, Agenda, Summary

        var dynamicSlides = slidePlan.SlideInstances.Where(s => s.Type == SlideInstanceType.Generated).ToList();
        dynamicSlides.Should().HaveCount(2); // 2 item slides
    }

    [Fact]
    public void GeneratePlan_EmptyCollection_CreatesEmptySlide()
    {
        // Arrange
        var generator = new SlidePlanGenerator();
        var slideInfos = new List<SlideInfo>
        {
            new SlideInfo
            {
                SlideId = 0,
                Type = SlideType.Source,
                Position = 0,
                CollectionName = "Items",
                MaxArrayIndex = 1
            }
        };

        var data = new Dictionary<string, object>
        {
            ["Items"] = new object[0] // Empty array
        };

        // Act
        var slidePlan = generator.GeneratePlan(slideInfos, data);

        // Assert
        slidePlan.Should().NotBeNull();
        slidePlan.SlideInstances.Should().HaveCount(1); // At least one slide should be created
        slidePlan.SlideInstances[0].IsEmpty.Should().BeTrue();
    }

    [Fact]
    public void GeneratePlan_AliasDirective_ResolvesCorrectly()
    {
        // Arrange
        var generator = new SlidePlanGenerator();
        var slideInfos = new List<SlideInfo>
        {
            new SlideInfo
            {
                SlideId = 0,
                Type = SlideType.Source,
                Position = 0,
                CollectionName = "Company.Departments>Employees",
                MaxArrayIndex = 1,
                Directives = new List<Directive>
                {
                    new Directive 
                    { 
                        Type = DirectiveType.Alias, 
                        SourcePath = "Company.Departments>Employees",
                        AliasName = "Staff"
                    }
                }
            }
        };

        var data = new Dictionary<string, object>
        {
            ["Company"] = new
            {
                Departments = new[]
                {
                    new
                    {
                        Name = "Sales",
                        Employees = new[]
                        {
                            new { Name = "John", Position = "Manager" },
                            new { Name = "Jane", Position = "Rep" }
                        }
                    }
                }
            }
        };

        // Act
        var slidePlan = generator.GeneratePlan(slideInfos, data);

        // Assert
        slidePlan.Should().NotBeNull();
        slidePlan.SlideInstances.Should().HaveCount(1);
        slidePlan.Aliases.Should().ContainKey("Staff");
        slidePlan.Aliases["Staff"].Should().Be("Company.Departments>Employees");
    }
}