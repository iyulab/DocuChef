using DocuChef.Presentation;
using DocuChef.Progress;
using DocuChef.Logging;

namespace DocuChef.TestConsoleApp;

public class ProgressTest
{
    public static void RunProgressDemo()
    {
        Console.WriteLine("=== PowerPoint Progress Demo ===");
        Console.WriteLine();

        var templatePath = Path.Combine("files", "ppt", "test-template.pptx");
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            Console.WriteLine("Please ensure test-template.pptx exists in the files/ppt directory.");
            return;
        }

        // 진행률 콜백 설정
        var progressCallback = new ProgressCallback(progress =>
        {
            var progressBar = CreateProgressBar(progress.OverallPercentage, 40);
            Console.Write($"\r{progressBar} {progress.OverallPercentage:D3}% - {progress.Message}");
            
            if (progress.OverallPercentage == 100)
            {
                Console.WriteLine(); // 완료 시 새 줄
            }
        });

        try
        {
            var options = new PowerPointOptions
            {
                EnableVerboseLogging = false
            };

            Console.WriteLine("Starting PowerPoint processing with progress tracking...");
            Console.WriteLine();

            using var recipe = new PowerPointRecipe(templatePath, options, progressCallback);
            
            // 테스트 데이터 추가
            recipe.AddVariable("title", "진행률 테스트 문서");
            recipe.AddVariable("subtitle", "DocuChef Progress Demo");
            recipe.AddVariable("name", "테스트 사용자");
            recipe.AddVariable("date", DateTime.Now.ToString("yyyy-MM-dd"));
            
            // 목록 데이터
            recipe.AddVariable("items", new[]
            {
                new { name = "항목 1", description = "첫 번째 테스트 항목" },
                new { name = "항목 2", description = "두 번째 테스트 항목" },
                new { name = "항목 3", description = "세 번째 테스트 항목" },
                new { name = "항목 4", description = "네 번째 테스트 항목" },
                new { name = "항목 5", description = "다섯 번째 테스트 항목" }
            });

            var outputPath = Path.Combine("output", "progress-test-output.pptx");
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            var result = recipe.Cook(outputPath);
            
            Console.WriteLine();
            Console.WriteLine($"✅ PowerPoint document generated successfully: {outputPath}");
            Console.WriteLine($"   File size: {new FileInfo(outputPath).Length:N0} bytes");
            
            result.Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine($"❌ Error: {ex.Message}");
            Console.WriteLine($"   Details: {ex}");
        }
    }

    public static void RunDetailedProgressDemo()
    {
        Console.WriteLine("=== Detailed Progress Demo ===");
        Console.WriteLine();

        var templatePath = Path.Combine("files", "ppt", "template_1.pptx");
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        var progressHistory = new List<ProcessingProgress>();
        
        var progressCallback = new ProgressCallback(progress =>
        {
            progressHistory.Add(new ProcessingProgress
            {
                Phase = progress.Phase,
                OverallPercentage = progress.OverallPercentage,
                PhasePercentage = progress.PhasePercentage,
                CurrentStep = progress.CurrentStep,
                TotalSteps = progress.TotalSteps,
                Message = progress.Message,
                Details = progress.Details
            });

            // 실시간 콘솔 출력
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {progress.Phase} - {progress.OverallPercentage}% - {progress.Message}");
            if (!string.IsNullOrEmpty(progress.Details))
            {
                Console.WriteLine($"    └─ {progress.Details}");
            }
        });

        try
        {
            var options = new PowerPointOptions
            {
                EnableVerboseLogging = true
            };

            using var recipe = new PowerPointRecipe(templatePath, options, progressCallback);
            
            recipe.AddVariable("company", "DocuChef Technologies");
            recipe.AddVariable("products", new[]
            {
                new { name = "PowerPoint 엔진", version = "1.0", status = "Active" },
                new { name = "Excel 엔진", version = "1.2", status = "Active" },
                new { name = "Word 엔진", version = "0.9", status = "Beta" }
            });

            var outputPath = Path.Combine("output", "detailed-progress-output.pptx");
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            var result = recipe.Cook(outputPath);
            
            Console.WriteLine();
            Console.WriteLine("=== Progress Summary ===");
            Console.WriteLine($"Total progress updates: {progressHistory.Count}");
            Console.WriteLine($"Processing phases covered: {progressHistory.Select(p => p.Phase).Distinct().Count()}");
            Console.WriteLine($"Total processing time: {DateTime.Now:HH:mm:ss}");
            
            // 각 단계별 요약
            var phaseGroups = progressHistory.GroupBy(p => p.Phase);
            foreach (var group in phaseGroups)
            {
                var updates = group.ToList();
                Console.WriteLine($"  {group.Key}: {updates.Count} updates ({updates.First().OverallPercentage}% → {updates.Last().OverallPercentage}%)");
            }
            
            result.Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
        }
    }

    private static string CreateProgressBar(int percentage, int width)
    {
        var filled = (int)(percentage * width / 100.0);
        var empty = width - filled;
        
        return $"[{new string('█', filled)}{new string('░', empty)}]";
    }
}