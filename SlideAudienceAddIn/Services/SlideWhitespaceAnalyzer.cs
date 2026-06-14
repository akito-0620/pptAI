using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using SlideAudienceAddIn.Models;

namespace SlideAudienceAddIn.Services
{
    public class WhitespaceRegion
    {
        public double X { get; set; }

        public double Y { get; set; }

        public double Width { get; set; }

        public double Height { get; set; }

        public double Score { get; set; }
    }

    public class SlideWhitespaceAnalysisResult
    {
        public List<WhitespaceRegion> Regions { get; set; } = new List<WhitespaceRegion>();

        public bool Succeeded { get; set; }

        public string FallbackReason { get; set; }
    }

    public class SlideWhitespaceAnalyzer
    {
        private readonly AppSettings _settings;

        public SlideWhitespaceAnalyzer(AppSettings settings)
        {
            _settings = settings ?? new AppSettings();
        }

        public SlideWhitespaceAnalysisResult Analyze(string imagePath)
        {
            Debug.WriteLine("[SlideAudience] Whitespace analysis started");

            try
            {
                if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                {
                    return Fail("image path missing");
                }

                using (var original = new Bitmap(imagePath))
                using (var bitmap = CreateAnalysisBitmap(original))
                {
                    Debug.WriteLine($"[SlideAudience] Whitespace image width={original.Width}, height={original.Height}");
                    Debug.WriteLine($"[SlideAudience] Whitespace analysis bitmap width={bitmap.Width}, height={bitmap.Height}");
                    var columns = Math.Max(4, _settings.Overlay.WhitespaceGridColumns);
                    var rows = Math.Max(3, _settings.Overlay.WhitespaceGridRows);
                    Debug.WriteLine($"[SlideAudience] Whitespace grid columns={columns}, rows={rows}");

                    var regions = new List<WhitespaceRegion>();
                    for (var row = 0; row < rows; row++)
                    {
                        for (var column = 0; column < columns; column++)
                        {
                            var score = ScoreCell(bitmap, column, row, columns, rows);
                            if (score >= 0.48)
                            {
                                regions.Add(new WhitespaceRegion
                                {
                                    X = (double)column / columns,
                                    Y = (double)row / rows,
                                    Width = 1.0 / columns,
                                    Height = 1.0 / rows,
                                    Score = Math.Round(score, 3)
                                });
                            }
                        }
                    }

                    regions = regions
                        .OrderByDescending(region => region.Score)
                        .Take(24)
                        .ToList();

                    Debug.WriteLine($"[SlideAudience] detected whitespace region count={regions.Count}");
                    foreach (var region in regions)
                    {
                        Debug.WriteLine(
                            $"[SlideAudience] whitespace region x={region.X:F3}, y={region.Y:F3}, width={region.Width:F3}, height={region.Height:F3}, score={region.Score:F3}");
                    }

                    return new SlideWhitespaceAnalysisResult
                    {
                        Regions = regions,
                        Succeeded = regions.Count > 0,
                        FallbackReason = regions.Count > 0 ? null : "no whitespace region detected"
                    };
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] Whitespace analysis failed: {ex.Message}");
                Debug.WriteLine(ex.ToString());
                return Fail(ex.Message);
            }
        }

        private static SlideWhitespaceAnalysisResult Fail(string reason)
        {
            Debug.WriteLine($"[SlideAudience] placement fallback reason={reason}");
            return new SlideWhitespaceAnalysisResult
            {
                Succeeded = false,
                FallbackReason = reason,
                Regions = new List<WhitespaceRegion>()
            };
        }

        private static Bitmap CreateAnalysisBitmap(Bitmap original)
        {
            const int maxWidth = 640;
            if (original.Width <= maxWidth)
            {
                return new Bitmap(original);
            }

            var scale = (double)maxWidth / original.Width;
            var height = Math.Max(1, (int)Math.Round(original.Height * scale));
            return new Bitmap(original, new Size(maxWidth, height));
        }

        private static double ScoreCell(Bitmap bitmap, int column, int row, int columns, int rows)
        {
            var left = column * bitmap.Width / columns;
            var right = Math.Max(left + 1, (column + 1) * bitmap.Width / columns);
            var top = row * bitmap.Height / rows;
            var bottom = Math.Max(top + 1, (row + 1) * bitmap.Height / rows);
            var sampleStepX = Math.Max(1, (right - left) / 18);
            var sampleStepY = Math.Max(1, (bottom - top) / 12);
            var values = new List<double>();
            var nearWhiteCount = 0;
            var darkCount = 0;
            var edgeCount = 0;
            var sampleCount = 0;

            for (var y = top; y < bottom; y += sampleStepY)
            {
                for (var x = left; x < right; x += sampleStepX)
                {
                    var brightness = Brightness(bitmap.GetPixel(x, y));
                    values.Add(brightness);
                    sampleCount++;
                    if (brightness > 222)
                    {
                        nearWhiteCount++;
                    }

                    if (brightness < 95)
                    {
                        darkCount++;
                    }

                    var x2 = Math.Min(bitmap.Width - 1, x + sampleStepX);
                    var y2 = Math.Min(bitmap.Height - 1, y + sampleStepY);
                    var dx = Math.Abs(brightness - Brightness(bitmap.GetPixel(x2, y)));
                    var dy = Math.Abs(brightness - Brightness(bitmap.GetPixel(x, y2)));
                    if (dx + dy > 46)
                    {
                        edgeCount++;
                    }
                }
            }

            if (sampleCount == 0)
            {
                return 0;
            }

            var average = values.Average();
            var variance = values.Sum(value => Math.Pow(value - average, 2)) / values.Count;
            var nearWhiteRatio = (double)nearWhiteCount / sampleCount;
            var darkRatio = (double)darkCount / sampleCount;
            var edgeRatio = (double)edgeCount / sampleCount;
            var calmRatio = Math.Max(0, 1.0 - variance / 4200.0);
            var edgeCalmRatio = Math.Max(0, 1.0 - edgeRatio * 2.2);
            var darkPenalty = Math.Min(0.45, darkRatio * 1.4);
            var outerBonus = column == 0 || row == 0 || column == columns - 1 || row == rows - 1 ? 0.08 : 0;

            return Clamp(nearWhiteRatio * 0.35 + calmRatio * 0.32 + edgeCalmRatio * 0.25 + outerBonus - darkPenalty, 0, 1);
        }

        private static double Brightness(Color color)
        {
            return color.R * 0.299 + color.G * 0.587 + color.B * 0.114;
        }

        private static double Clamp(double value, double min, double max)
        {
            return Math.Max(min, Math.Min(max, value));
        }
    }

    public class SlidePreloadService
    {
        private readonly object _sync = new object();
        private readonly Dictionary<string, SlidePreloadResult> _cache = new Dictionary<string, SlidePreloadResult>();
        private readonly Dictionary<string, Task<SlidePreloadResult>> _inflight = new Dictionary<string, Task<SlidePreloadResult>>();
        private readonly SlideWhitespaceAnalyzer _slideWhitespaceAnalyzer;
        private readonly CommentGenerationService _commentGenerationService;

        public SlidePreloadService(
            SlideWhitespaceAnalyzer slideWhitespaceAnalyzer,
            CommentGenerationService commentGenerationService)
        {
            _slideWhitespaceAnalyzer = slideWhitespaceAnalyzer;
            _commentGenerationService = commentGenerationService;
        }

        public bool TryGetCached(string cacheKey, out SlidePreloadResult result)
        {
            result = null;
            if (string.IsNullOrWhiteSpace(cacheKey))
            {
                return false;
            }

            lock (_sync)
            {
                return _cache.TryGetValue(cacheKey, out result);
            }
        }

        public Task<SlidePreloadResult> PreloadSnapshotAsync(
            SlideSnapshot snapshot,
            AppSettings settings,
            CancellationToken cancellationToken)
        {
            if (snapshot == null || string.IsNullOrWhiteSpace(snapshot.CacheKey))
            {
                return Task.FromResult(new SlidePreloadResult
                {
                    Succeeded = false,
                    ErrorMessage = "snapshot missing",
                    CreatedAtUtc = DateTime.UtcNow
                });
            }

            lock (_sync)
            {
                if (_cache.TryGetValue(snapshot.CacheKey, out var cached))
                {
                    Debug.WriteLine($"[SlideAudience] Preload cache already exists slideIndex={snapshot.SlideIndex}");
                    return Task.FromResult(cached);
                }

                if (_inflight.TryGetValue(snapshot.CacheKey, out var existingTask))
                {
                    Debug.WriteLine($"[SlideAudience] Preload already inflight slideIndex={snapshot.SlideIndex}");
                    return existingTask;
                }

                var task = Task.Run(() => PreloadCoreAsync(snapshot, settings, cancellationToken), cancellationToken);
                _inflight[snapshot.CacheKey] = task;
                return task;
            }
        }

        public void Clear()
        {
            lock (_sync)
            {
                _cache.Clear();
                _inflight.Clear();
            }

            Debug.WriteLine("[SlideAudience] Preload cache cleared");
        }

        public void TrimCache(int maxSlides)
        {
            lock (_sync)
            {
                TrimCacheLocked(maxSlides);
            }
        }

        public void InvalidateBySettingsChange()
        {
            Clear();
        }

        private async Task<SlidePreloadResult> PreloadCoreAsync(
            SlideSnapshot snapshot,
            AppSettings settings,
            CancellationToken cancellationToken)
        {
            var totalStopwatch = Stopwatch.StartNew();
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                Debug.WriteLine($"[SlideAudience] Preload started slideIndex={snapshot.SlideIndex}, slideId={snapshot.SlideId}");

                var whitespaceStopwatch = Stopwatch.StartNew();
                var whitespaceAnalysis = _slideWhitespaceAnalyzer?.Analyze(snapshot.ImagePath);
                whitespaceStopwatch.Stop();
                Debug.WriteLine($"[SlideAudience] Preload whitespace analysis took {whitespaceStopwatch.ElapsedMilliseconds} ms, slideIndex={snapshot.SlideIndex}");

                cancellationToken.ThrowIfCancellationRequested();

                var commentStopwatch = Stopwatch.StartNew();
                var commentResult = await _commentGenerationService.GenerateForSlideAsync(
                    snapshot.SlideId,
                    snapshot.SlideIndex,
                    snapshot.ImagePath,
                    snapshot.SlideText,
                    cancellationToken);
                commentStopwatch.Stop();
                cancellationToken.ThrowIfCancellationRequested();

                totalStopwatch.Stop();
                var result = new SlidePreloadResult
                {
                    CacheKey = snapshot.CacheKey,
                    SlideIndex = snapshot.SlideIndex,
                    SlideId = snapshot.SlideId,
                    ImagePath = snapshot.ImagePath,
                    SlideText = snapshot.SlideText,
                    WhitespaceAnalysis = whitespaceAnalysis,
                    CommentResult = commentResult,
                    UsedApi = commentResult != null && commentResult.UsedApi,
                    Succeeded = true,
                    SlideExportTimeMs = snapshot.SlideExportTimeMs,
                    TextExtractTimeMs = snapshot.TextExtractTimeMs,
                    WhitespaceAnalysisTimeMs = whitespaceStopwatch.ElapsedMilliseconds,
                    CommentGenerationTimeMs = commentStopwatch.ElapsedMilliseconds,
                    TotalPreloadTimeMs = totalStopwatch.ElapsedMilliseconds,
                    CreatedAtUtc = DateTime.UtcNow
                };

                lock (_sync)
                {
                    _cache[snapshot.CacheKey] = result;
                    TrimCacheLocked(settings?.Preload?.PreloadCacheMaxSlides ?? 5);
                }

                Debug.WriteLine($"[SlideAudience] Preload completed slideIndex={snapshot.SlideIndex}, totalMs={result.TotalPreloadTimeMs}, usedApi={result.UsedApi}");
                return result;
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine($"[SlideAudience] Preload cancelled slideIndex={snapshot.SlideIndex}");
                throw;
            }
            catch (Exception ex)
            {
                totalStopwatch.Stop();
                Debug.WriteLine($"[SlideAudience] Preload failed slideIndex={snapshot.SlideIndex}: {ex.Message}");
                Debug.WriteLine(ex.ToString());
                return new SlidePreloadResult
                {
                    CacheKey = snapshot.CacheKey,
                    SlideIndex = snapshot.SlideIndex,
                    SlideId = snapshot.SlideId,
                    ImagePath = snapshot.ImagePath,
                    SlideText = snapshot.SlideText,
                    Succeeded = false,
                    ErrorMessage = ex.Message,
                    SlideExportTimeMs = snapshot.SlideExportTimeMs,
                    TextExtractTimeMs = snapshot.TextExtractTimeMs,
                    TotalPreloadTimeMs = totalStopwatch.ElapsedMilliseconds,
                    CreatedAtUtc = DateTime.UtcNow
                };
            }
            finally
            {
                lock (_sync)
                {
                    _inflight.Remove(snapshot.CacheKey);
                }
            }
        }

        private void TrimCacheLocked(int maxSlides)
        {
            maxSlides = Math.Max(1, maxSlides);
            if (_cache.Count <= maxSlides)
            {
                return;
            }

            var removeKeys = _cache
                .OrderBy(item => item.Value.CreatedAtUtc)
                .Take(_cache.Count - maxSlides)
                .Select(item => item.Key)
                .ToList();

            foreach (var key in removeKeys)
            {
                _cache.Remove(key);
            }

            Debug.WriteLine($"[SlideAudience] Preload cache trimmed removed={removeKeys.Count}, remaining={_cache.Count}");
        }
    }
}
