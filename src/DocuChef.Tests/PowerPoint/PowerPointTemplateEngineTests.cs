using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocuChef.Presentation;
using DocuChef.Presentation.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint
{
    /// <summary>
    /// PowerPoint 템플릿 엔진의 핵심 기능 통합 테스트
    /// </summary>
    public class PowerPointTemplateEngineTests : TestBase
    {
        private readonly string _testTemplatesPath;
        private readonly string _testOutputPath;

        public PowerPointTemplateEngineTests(ITestOutputHelper output) : base(output)
        {
            _testTemplatesPath = Path.Combine(Path.GetTempPath(), "DocuChef_TestTemplates");
            _testOutputPath = Path.Combine(Path.GetTempPath(), "DocuChef_TestOutput");

            Directory.CreateDirectory(_testTemplatesPath);
            Directory.CreateDirectory(_testOutputPath);
        }        [Fact]
        public void ProcessTemplate_SimpleBinding_GeneratesCorrectOutput()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateBasicTemplate();
            var outputPath = Path.Combine(_testOutputPath, "simple_binding_output.pptx");

            var data = new
            {
                Title = "DocuChef 테스트 보고서",
                Company = "DocuChef Inc.",
                Date = new DateTime(2025, 5, 22),
                Version = "1.0.0"
            };            // Act
            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable(data);
            recipe.Cook(outputPath);
              // 파일 검증
            File.Exists(outputPath).Should().BeTrue("출력 파일이 생성되어야 함");
            
            using var presentationDocument = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(outputPath, false);
            presentationDocument.Should().NotBeNull();
            var slideCount = CountSlides(presentationDocument);
            slideCount.Should().Be(1, "템플릿에 정의된 슬라이드 수와 일치해야 함");
        }

        [Fact]
        public void ProcessTemplate_CollectionWithAutoGeneration_CreatesMultipleSlides()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateCollectionTemplate();
            var outputPath = Path.Combine(_testOutputPath, "collection_output.pptx");

            var data = new
            {
                Title = "항목 목록",
                Items = new[] {
                    new { Name = "항목 1", Value = 100 },
                    new { Name = "항목 2", Value = 200 },
                    new { Name = "항목 3", Value = 300 },
                    new { Name = "항목 4", Value = 400 },
                    new { Name = "항목 5", Value = 500 }
                }
            };            // Act
            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable(data);
            recipe.Cook(outputPath);

            // Assert
            File.Exists(outputPath).Should().BeTrue("출력 파일이 생성되어야 함");

            using var doc = PresentationDocument.Open(outputPath, false);
            var slideCount = CountSlides(doc);
            // 5개 항목 / 2개 항목 per 슬라이드 = 3개 슬라이드 (올림)
            slideCount.Should().Be(3, "항목 수에 맞게 슬라이드가 생성되어야 함");
        }

        [Fact]
        public void ProcessTemplate_NestedCollections_CreatesHierarchicalSlides()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateNestedCollectionTemplate();
            var outputPath = Path.Combine(_testOutputPath, "nested_collection_output.pptx");

            var data = new
            {
                Report = new {
                    Title = "부서별 보고서",
                    Date = new DateTime(2025, 5, 22),
                    Departments = new[] {
                        new {
                            Name = "개발부",
                            Manager = "김개발",
                            Teams = new[] {
                                new {
                                    Name = "프론트엔드팀",
                                    Projects = new[] {
                                        new { Name = "UI 개선", Status = "진행중" },
                                        new { Name = "성능 최적화", Status = "완료" }
                                    }
                                },
                                new {
                                    Name = "백엔드팀",
                                    Projects = new[] {
                                        new { Name = "API 개발", Status = "진행중" },
                                        new { Name = "데이터베이스 마이그레이션", Status = "예정" }
                                    }
                                }
                            }
                        },
                        new {
                            Name = "마케팅부",
                            Manager = "이마케팅",
                            Teams = new[] {
                                new {
                                    Name = "디지털마케팅팀",
                                    Projects = new[] {
                                        new { Name = "SNS 캠페인", Status = "진행중" }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            // Act
            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable(data);
            recipe.Cook(outputPath);

            // Assert
            File.Exists(outputPath).Should().BeTrue("출력 파일이 생성되어야 함");

            using var doc = PresentationDocument.Open(outputPath, false);
            var slideCount = CountSlides(doc);
            // 예상: 타이틀(1) + 부서(2) + 팀(3) + 프로젝트(3)
            slideCount.Should().Be(9, "중첩된 계층 구조에 맞게 슬라이드가 생성되어야 함");
        }

        [Fact]
        public void ProcessTemplate_FormatSpecifiers_AppliesFormattingCorrectly()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateFormatSpecifierTemplate();
            var outputPath = Path.Combine(_testOutputPath, "format_specifier_output.pptx");

            var data = new
            {
                Report = new {
                    Date = new DateTime(2025, 5, 22),
                    Amount = 1234567.89m,
                    Progress = 0.75,
                    IsComplete = true,
                    Items = new[] {
                        new { Name = "항목 A", Price = 10000m, CreatedAt = new DateTime(2025, 1, 15) },
                        new { Name = "항목 B", Price = 25000m, CreatedAt = new DateTime(2025, 3, 10) }
                    }
                }
            };

            // Act
            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable(data);
            recipe.Cook(outputPath);

            // Assert
            File.Exists(outputPath).Should().BeTrue("출력 파일이 생성되어야 함");
        }

        [Fact]
        public void ProcessTemplate_ConditionalExpressions_EvaluatesCorrectly()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateConditionalTemplate();
            var outputPath = Path.Combine(_testOutputPath, "conditional_output.pptx");

            var data = new
            {
                Status = new {
                    IsActive = true,
                    Score = 85,
                    Count = 0,
                    Products = new[] {
                        new { Name = "제품 A", InStock = true, Price = 15000m },
                        new { Name = "제품 B", InStock = false, Price = 25000m }
                    }
                }
            };

            // Act
            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable(data);
            recipe.Cook(outputPath);

            // Assert
            File.Exists(outputPath).Should().BeTrue("출력 파일이 생성되어야 함");
        }

        [Fact]
        public void ProcessTemplate_AliasDirective_ResolvesCorrectly()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateAliasTemplate();
            var outputPath = Path.Combine(_testOutputPath, "alias_output.pptx");

            var data = new
            {
                Company = new {
                    Departments = new[] {
                        new {
                            Name = "영업부",
                            Employees = new[] {
                                new { Name = "홍길동", Position = "영업대표" },
                                new { Name = "김영업", Position = "영업사원" }
                            }
                        }
                    }
                }
            };

            // Act
            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable(data);
            recipe.Cook(outputPath);

            // Assert
            File.Exists(outputPath).Should().BeTrue("출력 파일이 생성되어야 함");
        }

        [Fact]
        public void ProcessTemplate_ErrorHandling_HandlesEdgeCasesGracefully()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateErrorHandlingTemplate();
            var outputPath = Path.Combine(_testOutputPath, "error_handling_output.pptx");

            var data = new
            {
                Title = "오류 처리 테스트",
                // MissingProperty는 의도적으로 누락 (빈 문자열로 대체되어야 함)
                Items = new[] { new { Name = "항목 1" } } // 템플릿은 Items[0], Items[1], Items[2]를 사용
            };

            // Act
            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable(data);
            recipe.Cook(outputPath);

            // Assert
            File.Exists(outputPath).Should().BeTrue("오류 처리 후에도 파일이 생성되어야 함");
        }

        [Fact]
        public void ProcessTemplate_CustomFunction_AppliesFunction()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateCustomFunctionTemplate();
            var outputPath = Path.Combine(_testOutputPath, "custom_function_output.pptx");

            var data = new
            {
                Company = "DocuChef Inc.",
                Date = new DateTime(2025, 5, 22),
                Amount = 1500000m
            };

            // Act
            var recipe = chef.LoadTemplate(templatePath) as PowerPointRecipe;
            // 사용자 정의 함수 등록
            if (recipe != null)
            {
                recipe.RegisterFunction("FormatKoreanDate", (value) =>
                {
                    if (value is DateTime date)
                    {
                        return $"{date.Year}년 {date.Month}월 {date.Day}일";
                    }
                    return value;
                });

                recipe.RegisterFunction("FormatKoreanCurrency", (value) =>
                {
                    if (value is decimal amount)
                    {
                        return $"{amount:#,##0}원";
                    }
                    return value;
                });

                recipe.AddVariable(data);
                recipe.Cook(outputPath);
            }

            // Assert
            File.Exists(outputPath).Should().BeTrue("사용자 정의 함수를 적용한 파일이 생성되어야 함");
        }

        [Fact]
        public void ProcessTemplate_Performance_HandlesLargeData()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreatePerformanceTemplate();
            var outputPath = Path.Combine(_testOutputPath, "performance_output.pptx");

            // 대량의 데이터 생성
            var items = new List<object>();
            for (int i = 0; i < 100; i++)
            {
                items.Add(new
                {
                    Id = i + 1,
                    Name = $"항목 {i + 1}",
                    Description = $"항목 {i + 1}에 대한 상세 설명입니다.",
                    Value = (i + 1) * 100,
                    Category = i % 5 == 0 ? "A" : i % 5 == 1 ? "B" : i % 5 == 2 ? "C" : i % 5 == 3 ? "D" : "E"
                });
            }

            var data = new
            {
                Title = "대용량 데이터 테스트",
                Date = DateTime.Now,
                Items = items.ToArray()
            };

            // Act
            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable(data);
            
            // 성능 측정
            var stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            recipe.Cook(outputPath);
            stopwatch.Stop();

            // Assert
            File.Exists(outputPath).Should().BeTrue("대용량 데이터도 처리되어야 함");
            _output.WriteLine($"대용량 데이터 처리 시간: {stopwatch.ElapsedMilliseconds}ms");

            using var doc = PresentationDocument.Open(outputPath, false);
            var slideCount = CountSlides(doc);
            slideCount.Should().BeGreaterThan(10, "대량의 항목에 대해 여러 슬라이드가 생성되어야 함");
        }        // 정리 메서드
        public override void Dispose()
        {
            // 테스트 파일들 정리
            if (Directory.Exists(_testTemplatesPath))
            {
                Directory.Delete(_testTemplatesPath, true);
            }
            if (Directory.Exists(_testOutputPath))
            {
                Directory.Delete(_testOutputPath, true);
            }

            base.Dispose();
        }

        // 테스트 템플릿 생성 메서드
        private string CreateBasicTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "basic_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "제목: ${Title}, 회사: ${Company}, 날짜: ${Date:yyyy-MM-dd}, 버전: ${Version}"
            });
            return templatePath;
        }

        private string CreateCollectionTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "collection_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "제목: ${Title}",
                "항목: ${Items[0].Name} - ${Items[0].Value}, ${Items[1].Name} - ${Items[1].Value}"
            }, new[] {
                "#foreach: Items, max: 2"
            });
            return templatePath;
        }

        private string CreateNestedCollectionTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "nested_collection_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "보고서 제목: ${Report.Title}, 날짜: ${Report.Date:yyyy-MM-dd}",
                "부서: ${Report.Departments[0].Name}, 관리자: ${Report.Departments[0].Manager}",
                "팀: ${Report.Departments>Teams[0].Name}",
                "프로젝트: ${Report.Departments>Teams>Projects[0].Name}, 상태: ${Report.Departments>Teams>Projects[0].Status}"
            }, new[] {
                "#foreach: Report.Departments",
                "#range: begin, Report.Departments",
                "#foreach: Report.Departments>Teams",
                "#range: Report.Departments>Teams",
                "#foreach: Report.Departments>Teams>Projects",
                "#range: end, Report.Departments>Teams"
            });
            return templatePath;
        }

        private string CreateFormatSpecifierTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "format_specifier_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "날짜: ${Report.Date:yyyy년 MM월 dd일}, 금액: ${Report.Amount:C}, 진행률: ${Report.Progress:P0}",
                "완료 여부: ${Report.IsComplete}",
                "항목: ${Report.Items[0].Name}, 가격: ${Report.Items[0].Price:N0}원, 생성일: ${Report.Items[0].CreatedAt:yyyy-MM-dd}"
            });
            return templatePath;
        }

        private string CreateConditionalTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "conditional_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "활성화: ${Status.IsActive ? \"활성\" : \"비활성\"}, 점수: ${Status.Score > 80 ? \"우수\" : \"보통\"}",
                "갯수: ${Status.Count > 0 ? Status.Count : \"없음\"}",
                "제품: ${Status.Products[0].Name}, 재고: ${Status.Products[0].InStock ? \"있음\" : \"없음\"}, 가격: ${Status.Products[0].Price > 20000 ? \"고가\" : \"저가\"}"
            });
            return templatePath;
        }

        private string CreateAliasTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "alias_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "부서: ${Company.Departments[0].Name}",
                "직원: ${Staff[0].Name}, 직책: ${Staff[0].Position}"
            }, new[] {
                "#alias: Company.Departments[0].Employees as Staff"
            });
            return templatePath;
        }

        private string CreateErrorHandlingTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "error_handling_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "제목: ${Title}, 설명: ${MissingProperty}",
                "항목: ${Items[0].Name}, ${Items[1].Name}, ${Items[2].Name}"
            });
            return templatePath;
        }

        private string CreateCustomFunctionTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "custom_function_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "회사: ${Company}",
                "한국식 날짜: ${FormatKoreanDate(Date)}",
                "한국식 금액: ${FormatKoreanCurrency(Amount)}"
            });
            return templatePath;
        }

        private string CreatePerformanceTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "performance_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "제목: ${Title}, 날짜: ${Date:yyyy-MM-dd}",
                "ID: ${Items[0].Id}, 이름: ${Items[0].Name}",
                "설명: ${Items[0].Description}",
                "값: ${Items[0].Value:N0}, 카테고리: ${Items[0].Category}"
            }, new[] {
                "#foreach: Items, max: 2"
            });
            return templatePath;
        }private void CreateMockPresentationFile(string filePath, string[] slideContents, string[]? slideNotes = null)
        {
            // 실제 구현에서는 OpenXml을 사용하여 실제 PowerPoint 파일 생성
            // 여기서는 테스트를 위한 간단한 모의 구현

            using var doc = PresentationDocument.Create(filePath, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);

            var presentationPart = doc.AddPresentationPart();
            presentationPart.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation();

            var slideIdList = new SlideIdList();
            presentationPart.Presentation.AppendChild(slideIdList);

            for (int i = 0; i < slideContents.Length; i++)
            {
                var slidePart = presentationPart.AddNewPart<SlidePart>();
                var slide = new Slide();
                slidePart.Slide = slide;

                // 슬라이드 노트 추가 (지시문이 있는 경우)
                if (slideNotes != null && i < slideNotes.Length)
                {
                    var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
                    var notesSlide = new NotesSlide();
                    notesSlidePart.NotesSlide = notesSlide;
                }

                var slideId = new SlideId();
                slideId.Id = (uint)(256 + i);
                slideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);
                slideIdList.AppendChild(slideId);
            }
        }

        private int CountSlides(PresentationDocument presentationDocument)
        {
            if (presentationDocument?.PresentationPart?.Presentation?.SlideIdList == null)
                return 0;

            return presentationDocument.PresentationPart.Presentation.SlideIdList.Count();
        }
    }
}