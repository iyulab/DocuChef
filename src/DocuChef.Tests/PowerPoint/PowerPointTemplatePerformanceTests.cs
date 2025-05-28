using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace DocuChef.Tests.PowerPoint
{
    /// <summary>
    /// 성능 테스트 클래스
    /// </summary>
    public class PowerPointTemplatePerformanceTests : TestBase
    {
        private readonly string _testTemplatesPath;
        private readonly string _testOutputPath;

        public PowerPointTemplatePerformanceTests(ITestOutputHelper output) : base(output)
        {
            _testTemplatesPath = Path.Combine(Path.GetTempPath(), "DocuChef_PerformanceTestTemplates");
            _testOutputPath = Path.Combine(Path.GetTempPath(), "DocuChef_PerformanceTestOutput");

            Directory.CreateDirectory(_testTemplatesPath);
            Directory.CreateDirectory(_testOutputPath);
        }

        [Fact]
        public void ProcessLargeTemplate_WithManySlides_CompletesWithinTimeLimit()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateLargeTemplate();
            var outputPath = Path.Combine(_testOutputPath, "large_template_output.pptx");

            // 대량의 데이터 생성
            var products = new List<object>();
            for (int i = 0; i < 100; i++)
            {
                products.Add(new
                {
                    Id = i,
                    Name = $"제품 {i}",
                    Description = $"제품 {i}에 대한 상세한 설명입니다. 이 제품은 고품질의 재료로 제작되었으며 다양한 용도로 사용할 수 있습니다.",
                    Price = 1000m + (i * 100),
                    Category = i % 5 == 0 ? "전자제품" : i % 5 == 1 ? "가구" : i % 5 == 2 ? "의류" : i % 5 == 3 ? "도서" : "기타",
                    InStock = i % 3 != 0,
                    ReleaseDate = DateTime.Now.AddDays(-i)
                });
            }

            var data = new
            {
                Title = "대량 데이터 성능 테스트",
                Date = DateTime.Now,
                Company = "테스트 회사",
                Products = products.ToArray()
            };

            // Act
            var stopwatch = Stopwatch.StartNew();

            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable("Title", data.Title);
            recipe.AddVariable("Date", data.Date);
            recipe.AddVariable("Company", data.Company);
            recipe.AddVariable("Products", data.Products);
            recipe.Cook(outputPath);

            stopwatch.Stop();

            // Assert
            _output.WriteLine($"대량 데이터 처리 시간: {stopwatch.ElapsedMilliseconds}ms");

            // 5초 이내에 완료되어야 함 (성능 요구 사항에 따라 조정 가능)
            stopwatch.ElapsedMilliseconds.Should().BeLessThan(5000, "대량 데이터 처리는 5초 이내에 완료되어야 함");
            File.Exists(outputPath).Should().BeTrue("출력 파일이 생성되어야 함");

            // 생성된 슬라이드 수 확인
            using var doc = PresentationDocument.Open(outputPath, false);
            var slideCount = CountSlides(doc);
            // 타이틀 + (100개 제품 ÷ 슬라이드당 3개 제품 = 34개 슬라이드) (올림)
            slideCount.Should().Be(35, "예상된 수의 슬라이드가 생성되어야 함");
        }

        [Fact]
        public void ProcessDeepNestedData_CompletesSuccessfully()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateDeepNestedTemplate();
            var outputPath = Path.Combine(_testOutputPath, "deep_nested_output.pptx");

            // 깊은 중첩 데이터 생성 (4단계 깊이)
            var reports = new List<object>();
            for (int r = 0; r < 2; r++)
            {
                var departments = new List<object>();
                for (int d = 0; d < 3; d++)
                {
                    var teams = new List<object>();
                    for (int t = 0; t < 3; t++)
                    {
                        var members = new List<object>();
                        for (int m = 0; m < 5; m++)
                        {
                            members.Add(new
                            {
                                Name = $"멤버 {r}-{d}-{t}-{m}",
                                Position = $"직책 {m % 3}",
                                Salary = 50000 + (m * 10000),
                                Joined = DateTime.Now.AddMonths(-(m * 6)),
                                Skills = new[] { "C#", "JavaScript", "SQL" }.Take(m % 3 + 1).ToArray()
                            });
                        }

                        teams.Add(new
                        {
                            Name = $"팀 {r}-{d}-{t}",
                            Leader = $"팀장 {r}-{d}-{t}",
                            Budget = 1000000 + (t * 500000),
                            Members = members.ToArray()
                        });
                    }

                    departments.Add(new
                    {
                        Name = $"부서 {r}-{d}",
                        Manager = $"부서장 {r}-{d}",
                        Location = $"빌딩 {d + 1}",
                        Teams = teams.ToArray()
                    });
                }

                reports.Add(new
                {
                    Title = $"보고서 {r}",
                    Quarter = $"Q{r + 1}",
                    Year = 2025,
                    Departments = departments.ToArray()
                });
            }

            var data = new
            {
                Company = "테스트 회사",
                CEO = "홍길동",
                Reports = reports.ToArray()
            };

            // Act
            var stopwatch = Stopwatch.StartNew();

            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable("Company", data.Company);
            recipe.AddVariable("CEO", data.CEO);
            recipe.AddVariable("Reports", data.Reports);
            recipe.Cook(outputPath);

            stopwatch.Stop();

            // Assert
            _output.WriteLine($"깊은 중첩 데이터 처리 시간: {stopwatch.ElapsedMilliseconds}ms");

            // 처리 시간 확인 (깊은 중첩으로 인해 더 긴 시간 허용)
            stopwatch.ElapsedMilliseconds.Should().BeLessThan(10000, "깊은 중첩 데이터 처리는 10초 이내에 완료되어야 함");
            File.Exists(outputPath).Should().BeTrue("출력 파일이 생성되어야 함");

            using var doc = PresentationDocument.Open(outputPath, false);
            var slideCount = CountSlides(doc);
            slideCount.Should().BeGreaterThan(10, "중첩 데이터에 따라 많은 슬라이드가 생성되어야 함");
        }

        [Fact]
        public void ProcessTemplate_WithManyBindingExpressions_HandlesEfficiently()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateTemplateWithManyExpressions();
            var outputPath = Path.Combine(_testOutputPath, "many_expressions_output.pptx");

            // 많은 속성을 가진 데이터 객체 생성
            var data = new
            {
                Title = "바인딩 표현식 성능 테스트",
                Company = "테스트 회사",
                Department = "개발부",
                Manager = "홍길동",
                Date = DateTime.Now,
                Year = DateTime.Now.Year,
                Month = DateTime.Now.Month,
                Day = DateTime.Now.Day,
                Quarter = (DateTime.Now.Month - 1) / 3 + 1,
                Revenue = 10000000m,
                Expenses = 7500000m,
                Profit = 2500000m,
                GrowthRate = 15.5,
                EmployeeCount = 150,
                ProjectsCount = 12,
                CompletedProjects = 8,
                InProgressProjects = 4,
                CustomerSatisfaction = 4.8,
                Address = "서울시 강남구 테헤란로 123",
                Phone = "02-123-4567",
                Email = "info@test.com",
                Website = "www.test.com",
                CEO = "김대표",
                CTO = "이기술",
                CFO = "박재무",
                HR = "최인사",
                MarketCap = 50000000000m,
                Employees = Enumerable.Range(1, 20).Select(i => new {
                    Id = i,
                    Name = $"직원 {i}",
                    Position = $"직책 {i % 5}",
                    Salary = 50000 + (i * 5000),
                    Department = $"부서 {i % 3}",
                    JoinDate = DateTime.Now.AddYears(-i % 10),
                    Performance = (i % 5) * 20 + 20
                }).ToArray()
            };

            // Act
            var stopwatch = Stopwatch.StartNew();

            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable(data); // 전체 객체를 한 번에 추가
            recipe.Cook(outputPath);

            stopwatch.Stop();

            // Assert
            _output.WriteLine($"많은 바인딩 표현식 처리 시간: {stopwatch.ElapsedMilliseconds}ms");

            stopwatch.ElapsedMilliseconds.Should().BeLessThan(2000, "많은 바인딩 표현식 처리는 2초 이내에 완료되어야 함");
            File.Exists(outputPath).Should().BeTrue("출력 파일이 생성되어야 함");
        }

        [Fact]
        public void ProcessTemplate_ConcurrentProcessing_HandlesMultipleRequests()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateSampleTemplate();
            const int concurrentRequests = 5;
            var tasks = new List<Task<TimeSpan>>();

            // Act
            for (int i = 0; i < concurrentRequests; i++)
            {
                int requestId = i;
                var task = Task.Run(() => {
                    var stopwatch = Stopwatch.StartNew();

                    var outputPath = Path.Combine(_testOutputPath, $"concurrent_output_{requestId}.pptx");
                    var data = new
                    {
                        Title = $"동시 처리 테스트 {requestId}",
                        Items = Enumerable.Range(1, 10).Select(j => new {
                            Name = $"항목 {requestId}-{j}",
                            Value = j * 100
                        }).ToArray()
                    };

                    var recipe = chef.LoadTemplate(templatePath);
                    recipe.AddVariable("Title", data.Title);
                    recipe.AddVariable("Items", data.Items);
                    recipe.Cook(outputPath);

                    stopwatch.Stop();
                    return stopwatch.Elapsed;
                });

                tasks.Add(task);
            }

            // Wait for all tasks to complete
            var results = Task.WhenAll(tasks).GetAwaiter().GetResult();

            // Assert
            results.Should().HaveCount(concurrentRequests, "모든 동시 요청이 완료되어야 함");

            var maxProcessingTime = results.Max();
            var avgProcessingTime = results.Average(ts => ts.TotalMilliseconds);

            _output.WriteLine($"동시 처리 최대 시간: {maxProcessingTime.TotalMilliseconds}ms");
            _output.WriteLine($"동시 처리 평균 시간: {avgProcessingTime}ms");

            maxProcessingTime.TotalMilliseconds.Should().BeLessThan(5000, "동시 처리에서도 합리적인 시간 내에 완료되어야 함");

            // 모든 출력 파일이 생성되었는지 확인
            for (int i = 0; i < concurrentRequests; i++)
            {
                var outputPath = Path.Combine(_testOutputPath, $"concurrent_output_{i}.pptx");
                File.Exists(outputPath).Should().BeTrue($"출력 파일 {i}가 생성되어야 함");
            }
        }

        [Fact]
        public void ProcessTemplate_MemoryUsage_StaysWithinLimits()
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateLargeTemplate();
            var outputPath = Path.Combine(_testOutputPath, "memory_test_output.pptx");

            // 메모리 사용량 측정 시작
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            var initialMemory = GC.GetTotalMemory(false);

            // 중간 크기의 데이터 생성
            var data = new
            {
                Title = "메모리 사용량 테스트",
                Items = Enumerable.Range(1, 500).Select(i => new {
                    Id = i,
                    Name = $"항목 {i}",
                    Description = new string('A', 100), // 100자 문자열
                    Value = i * 1.5,
                    Category = $"카테고리 {i % 10}",
                    Tags = Enumerable.Range(1, 5).Select(j => $"태그{i}-{j}").ToArray()
                }).ToArray()
            };

            // Act
            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable("Title", data.Title);
            recipe.AddVariable("Items", data.Items);
            recipe.Cook(outputPath);

            // 메모리 사용량 측정 완료
            var finalMemory = GC.GetTotalMemory(true);
            var memoryUsed = finalMemory - initialMemory;

            // Assert
            _output.WriteLine($"메모리 사용량: {memoryUsed / 1024 / 1024:F2} MB");

            // 메모리 사용량이 100MB를 초과하지 않아야 함 (합리적인 제한)
            memoryUsed.Should().BeLessThan(100 * 1024 * 1024, "메모리 사용량이 100MB를 초과하지 않아야 함");
            File.Exists(outputPath).Should().BeTrue("출력 파일이 생성되어야 함");
        }

        [Theory]
        [InlineData(10, 1000)]    // 10개 항목 - 1초 이내
        [InlineData(50, 2000)]    // 50개 항목 - 2초 이내  
        [InlineData(100, 3000)]   // 100개 항목 - 3초 이내
        [InlineData(200, 5000)]   // 200개 항목 - 5초 이내
        public void ProcessTemplate_VariousDataSizes_CompletesWithinExpectedTime(int itemCount, int maxTimeMs)
        {
            // Arrange
            var chef = CreateNewChef();
            var templatePath = CreateSampleTemplate();
            var outputPath = Path.Combine(_testOutputPath, $"size_test_{itemCount}_output.pptx");

            var data = new
            {
                Title = $"크기별 성능 테스트 - {itemCount}개 항목",
                Items = Enumerable.Range(1, itemCount).Select(i => new {
                    Id = i,
                    Name = $"항목 {i}",
                    Value = i * 10
                }).ToArray()
            };

            // Act
            var stopwatch = Stopwatch.StartNew();

            var recipe = chef.LoadTemplate(templatePath);
            recipe.AddVariable("Title", data.Title);
            recipe.AddVariable("Items", data.Items);
            recipe.Cook(outputPath);

            stopwatch.Stop();

            // Assert
            _output.WriteLine($"{itemCount}개 항목 처리 시간: {stopwatch.ElapsedMilliseconds}ms");

            stopwatch.ElapsedMilliseconds.Should().BeLessThan(maxTimeMs,
                $"{itemCount}개 항목 처리는 {maxTimeMs}ms 이내에 완료되어야 함");
            File.Exists(outputPath).Should().BeTrue("출력 파일이 생성되어야 함");
        }

        // 정리 메서드
        public override void Dispose()
        {
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

        // 헬퍼 메서드들
        private string CreateLargeTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "large_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "제목: ${Title} - ${Date:yyyy-MM-dd}",
                "회사: ${Company}",
                "제품 목록: ${Products[0].Name}, ${Products[1].Name}, ${Products[2].Name}"
            });
            return templatePath;
        }

        private string CreateDeepNestedTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "deep_nested_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "회사: ${Company}, CEO: ${CEO}",
                "보고서: ${Reports[0].Title} - ${Reports[0].Year}",
                "부서: ${Reports>Departments[0].Name}",
                "팀: ${Reports>Departments>Teams[0].Name}",
                "멤버: ${Reports>Departments>Teams>Members[0].Name}, ${Reports>Departments>Teams>Members[1].Name}"
            });
            return templatePath;
        }

        private string CreateTemplateWithManyExpressions()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "many_expressions_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "기본 정보: ${Title}, ${Company}, ${Department}, ${Manager}",
                "날짜 정보: ${Date:yyyy-MM-dd}, ${Year}, ${Month}, ${Day}, Q${Quarter}",
                "재무 정보: ${Revenue:N0}, ${Expenses:N0}, ${Profit:N0}, ${GrowthRate:P2}",
                "인사 정보: ${EmployeeCount}명, ${ProjectsCount}개 프로젝트, 만족도 ${CustomerSatisfaction}",
                "연락처: ${Address}, ${Phone}, ${Email}, ${Website}",
                "임원진: CEO ${CEO}, CTO ${CTO}, CFO ${CFO}, HR ${HR}",
                "직원 목록: ${Employees[0].Name}, ${Employees[1].Name}, ${Employees[2].Name}"
            });
            return templatePath;
        }

        private string CreateSampleTemplate()
        {
            var templatePath = Path.Combine(_testTemplatesPath, "sample_template.pptx");
            CreateMockPresentationFile(templatePath, new[] {
                "제목: ${Title}",
                "항목: ${Items[0].Name}, ${Items[1].Name}"
            });
            return templatePath;
        }

        private void CreateMockPresentationFile(string filePath, string[] slideContents)
        {
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