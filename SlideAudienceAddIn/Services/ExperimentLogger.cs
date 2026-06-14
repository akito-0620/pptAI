using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using SlideAudienceAddIn.Models;
using SlideAudienceAddIn.Overlay;
using SlideAudienceAddIn.Utils;

namespace SlideAudienceAddIn.Services
{
    public class ExperimentLogger
    {
        private readonly AppSettings _settings;
        private readonly string _logDir;
        private readonly object _logLock = new object();
        private string _sessionId;
        private string _presentationHash;
        private string _logFilePath;

        public ExperimentLogger(AppSettings settings)
        {
            _settings = settings ?? new AppSettings();
            _logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "SlideAudience",
                "logs");
        }

        public void StartSession(string presentationPath)
        {
            _sessionId = Guid.NewGuid().ToString("N");
            _presentationHash = HashPath(presentationPath);
            _logFilePath = Path.Combine(_logDir, $"session_{DateTime.Now:yyyyMMdd_HHmmss}.jsonl");

            Log("session_started");
        }

        public void EndSession()
        {
            Log("session_ended");
        }

        public void LogSlideChanged(int slideId, int slideIndex)
        {
            Log("slide_changed", slideId, slideIndex);
        }

        public void LogSlideAnalyzed(
            int slideId,
            int slideIndex,
            string exportedImagePath,
            string extractedText,
            bool imageExportSucceeded,
            bool textExtractionSucceeded,
            SlideWhitespaceAnalysisResult whitespaceAnalysis,
            long? slideExportTimeMs = null,
            long? whitespaceAnalysisTimeMs = null)
        {
            var data = new Dictionary<string, object>
            {
                ["exportedImagePath"] = _settings.Experiment.SaveImagePathToLog ? exportedImagePath : null,
                ["extractedTextLength"] = string.IsNullOrEmpty(extractedText) ? 0 : extractedText.Length,
                ["imageExportSucceeded"] = imageExportSucceeded,
                ["textExtractionSucceeded"] = textExtractionSucceeded,
                ["slideExportTimeMs"] = slideExportTimeMs,
                ["whitespaceAnalysisTimeMs"] = whitespaceAnalysisTimeMs,
                ["whitespaceRegionCount"] = whitespaceAnalysis?.Regions?.Count ?? 0,
                ["whitespaceAnalysisSucceeded"] = whitespaceAnalysis != null && whitespaceAnalysis.Succeeded,
                ["whitespaceFallbackReason"] = whitespaceAnalysis?.FallbackReason
            };

            if (_settings.Experiment.SaveSlideTextToLog)
            {
                data["extractedSlideText"] = extractedText;
            }

            Log("slide_analyzed", slideId, slideIndex, data);
        }

        public void LogCommentsGenerated(int slideId, int slideIndex, CommentGenerationResult result)
        {
            var data = new Dictionary<string, object>
            {
                ["generatedComments"] = ToLogComments(result != null ? result.Comments : null),
                ["usedApi"] = result != null ? (object)result.UsedApi : null,
                ["latencyMs"] = result != null ? (object)result.LatencyMs : null,
                ["fallbackReason"] = result != null ? result.FallbackReason ?? result.ErrorMessage : null,
                ["model"] = result != null ? result.Model : _settings.Gemini.Model,
                ["promptLength"] = result != null ? (object)result.PromptLength : null,
                ["extractedTextLength"] = result != null ? (object)result.ExtractedTextLength : null
            };

            Log("comments_generated", slideId, slideIndex, data, result != null ? (bool?)result.UsedApi : null, result != null ? (long?)result.LatencyMs : null);
        }

        public void LogCommentsShown(
            int slideId,
            int slideIndex,
            IReadOnlyList<AudienceComment> comments,
            IReadOnlyList<WhitespaceRegion> whitespaceRegions)
        {
            var overlay = _settings.Overlay ?? new OverlaySettings();
            var regionCount = whitespaceRegions?.Count ?? 0;
            var data = new Dictionary<string, object>
            {
                ["shownComments"] = ToLogComments(comments, true, overlay),
                ["displayMode"] = overlay.DisplayMode.ToString(),
                ["overlayPosition"] = overlay.Position,
                ["commentLifetimeSeconds"] = overlay.CommentLifetimeSeconds,
                ["flowSpeedPixelsPerSecond"] = overlay.FlowSpeedPixelsPerSecond,
                ["bubbleLifetimeSeconds"] = overlay.BubbleLifetimeSeconds,
                ["bubbleFadeInSeconds"] = overlay.BubbleFadeInSeconds,
                ["bubbleFadeOutSeconds"] = overlay.BubbleFadeOutSeconds,
                ["commentDisplayIntervalMinSeconds"] = overlay.CommentDisplayIntervalMinSeconds,
                ["commentDisplayIntervalMaxSeconds"] = overlay.CommentDisplayIntervalMaxSeconds,
                ["scheduledDelaySeconds"] = null,
                ["whitespacePlacementProbability"] = overlay.WhitespacePlacementProbability,
                ["commentFontSize"] = overlay.CommentFontSize,
                ["commentTextColor"] = overlay.CommentTextColor,
                ["commentBackgroundColor"] = overlay.CommentBackgroundColor,
                ["whitespaceRegionCount"] = regionCount,
                ["selectedPlacementRegion"] = regionCount > 0 ? ToLogRegion(whitespaceRegions.OrderByDescending(region => region.Score).FirstOrDefault()) : null,
                ["placementMode"] = overlay.UseWhitespaceAwarePlacement && regionCount > 0 ? "whitespace" : "fallback",
                ["overlaySettings"] = ToOverlaySettingsSnapshot(overlay)
            };

            Log("comments_shown", slideId, slideIndex, data);
        }

        public void LogOverlayDiagnostics(int? slideId, int? slideIndex, string stage, OverlayBounds? bounds)
        {
            Log("overlay_diagnostics", slideId, slideIndex, new Dictionary<string, object>
            {
                ["stage"] = stage,
                ["overlayBounds"] = ToLogOverlayBounds(bounds),
                ["screens"] = ToLogScreens(),
                ["presentationMonitorDeviceName"] = _settings.Overlay.PresentationMonitorDeviceName,
                ["presentationMonitorIndex"] = _settings.Overlay.PresentationMonitorIndex
            });
        }

        public void LogCommentsHidden(int? slideId = null, int? slideIndex = null)
        {
            Log("comments_hidden", slideId, slideIndex);
        }

        public void LogOverlayModeChanged(string previousMode, string currentMode)
        {
            Log("overlay_mode_changed", null, null, new Dictionary<string, object>
            {
                ["previousDisplayMode"] = previousMode,
                ["currentDisplayMode"] = currentMode
            });
        }

        public void LogConditionPresetChanged(string preset)
        {
            Log("condition_preset_changed", null, null, new Dictionary<string, object>
            {
                ["conditionPreset"] = preset
            });
        }

        public void LogSettingsChanged(string settingName, object value)
        {
            Log("settings_changed", null, null, new Dictionary<string, object>
            {
                ["settingName"] = settingName,
                ["value"] = value,
                ["commentDisplayIntervalMinSeconds"] = _settings.Overlay.CommentDisplayIntervalMinSeconds,
                ["commentDisplayIntervalMaxSeconds"] = _settings.Overlay.CommentDisplayIntervalMaxSeconds,
                ["whitespacePlacementProbability"] = _settings.Overlay.WhitespacePlacementProbability,
                ["commentFontSize"] = _settings.Overlay.CommentFontSize,
                ["commentTextColor"] = _settings.Overlay.CommentTextColor,
                ["commentBackgroundColor"] = _settings.Overlay.CommentBackgroundColor,
                ["flowSpeedPixelsPerSecond"] = _settings.Overlay.FlowSpeedPixelsPerSecond,
                ["bubbleFadeInSeconds"] = _settings.Overlay.BubbleFadeInSeconds,
                ["bubbleFadeOutSeconds"] = _settings.Overlay.BubbleFadeOutSeconds,
                ["presentationMonitorDeviceName"] = _settings.Overlay.PresentationMonitorDeviceName,
                ["presentationMonitorIndex"] = _settings.Overlay.PresentationMonitorIndex
            });
        }

        public void LogPreloadStarted(SlideSnapshot snapshot)
        {
            Log("preload_started", snapshot?.SlideId, snapshot?.SlideIndex, ToPreloadData(snapshot, null, cacheHit: false));
        }

        public void LogPreloadCompleted(SlidePreloadResult result)
        {
            Log("preload_completed", result?.SlideId, result?.SlideIndex, ToPreloadData(null, result, cacheHit: false), result?.UsedApi, result?.CommentResult?.LatencyMs);
        }

        public void LogPreloadFailed(SlidePreloadResult result)
        {
            Log("preload_failed", result?.SlideId, result?.SlideIndex, ToPreloadData(null, result, cacheHit: false));
        }

        public void LogPreloadFailed(SlideSnapshot snapshot, Exception exception)
        {
            var data = ToPreloadData(snapshot, null, cacheHit: false);
            data["preloadSucceeded"] = false;
            data["errorMessage"] = exception?.GetBaseException()?.Message;
            Log("preload_failed", snapshot?.SlideId, snapshot?.SlideIndex, data);
        }

        public void LogPreloadUsed(SlidePreloadResult result, bool cacheHit)
        {
            Log("preload_used", result?.SlideId, result?.SlideIndex, ToPreloadData(null, result, cacheHit), result?.UsedApi, result?.CommentResult?.LatencyMs);
        }

        public void LogPreloadCacheMiss(int slideId, int slideIndex, string cacheKey)
        {
            Log("preload_cache_miss", slideId, slideIndex, new Dictionary<string, object>
            {
                ["cacheKeyHash"] = HashPath(cacheKey),
                ["cacheHit"] = false
            });
        }

        public void LogPreloadCancelled(SlideSnapshot snapshot)
        {
            var data = ToPreloadData(snapshot, null, cacheHit: false);
            data["preloadSucceeded"] = false;
            Log("preload_cancelled", snapshot?.SlideId, snapshot?.SlideIndex, data);
        }

        public void LogPreloadCacheCleared(string reason)
        {
            Log("preload_cache_cleared", null, null, new Dictionary<string, object>
            {
                ["reason"] = reason
            });
        }

        public void LogError(int? slideId, int? slideIndex, Exception exception)
        {
            Log("error", slideId, slideIndex, new Dictionary<string, object>
            {
                ["errorType"] = exception != null ? exception.GetType().FullName : null,
                ["message"] = exception != null ? exception.Message : null
            });
        }

        private void Log(
            string eventType,
            int? slideId = null,
            int? slideIndex = null,
            Dictionary<string, object> eventData = null,
            bool? usedApi = null,
            long? latencyMs = null)
        {
            if (!_settings.Experiment.EnableLogging)
            {
                return;
            }

            try
            {
                EnsureSession();
                Directory.CreateDirectory(_logDir);

                var entry = CreateCommonEntry(eventType, slideId, slideIndex, usedApi, latencyMs);
                if (eventData != null)
                {
                    foreach (var item in eventData)
                    {
                        entry[item.Key] = item.Value;
                    }
                }

                lock (_logLock)
                {
                    File.AppendAllText(_logFilePath, JsonHelper.Serialize(entry) + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SlideAudience] ExperimentLogger failed to write log");
                Debug.WriteLine(ex.ToString());
            }
        }

        private Dictionary<string, object> CreateCommonEntry(
            string eventType,
            int? slideId,
            int? slideIndex,
            bool? usedApi,
            long? latencyMs)
        {
            var overlay = _settings.Overlay ?? new OverlaySettings();
            return new Dictionary<string, object>
            {
                ["timestampUtc"] = DateTimeOffset.UtcNow.ToString("o"),
                ["eventType"] = eventType,
                ["sessionId"] = _sessionId,
                ["presentationHash"] = _presentationHash,
                ["slideId"] = slideId,
                ["slideIndex"] = slideIndex,
                ["displayMode"] = overlay.DisplayMode.ToString(),
                ["enableApi"] = _settings.Gemini.EnableApi,
                ["usedApi"] = usedApi,
                ["latencyMs"] = latencyMs
            };
        }

        private void EnsureSession()
        {
            if (string.IsNullOrWhiteSpace(_sessionId))
            {
                _sessionId = Guid.NewGuid().ToString("N");
            }

            if (string.IsNullOrWhiteSpace(_logFilePath))
            {
                _logFilePath = Path.Combine(_logDir, $"session_{DateTime.Now:yyyyMMdd_HHmmss}.jsonl");
            }
        }

        private static IReadOnlyList<Dictionary<string, object>> ToLogComments(
            IReadOnlyList<AudienceComment> comments,
            bool includeTiming = false,
            OverlaySettings overlay = null)
        {
            var start = DateTimeOffset.UtcNow;
            return (comments ?? new List<AudienceComment>())
                .Where(comment => comment != null)
                .Select((comment, index) =>
                {
                    var item = new Dictionary<string, object>
                    {
                        ["type"] = comment.Type,
                        ["persona"] = comment.Persona,
                        ["text"] = comment.Text,
                        ["length"] = string.IsNullOrEmpty(comment.Text) ? 0 : comment.Text.Length,
                        ["commentIndex"] = index,
                        ["displayOrder"] = index
                    };

                    if (includeTiming)
                    {
                        var maxSimultaneous = Math.Max(1, Math.Min(3, overlay?.MaxSimultaneousComments ?? 3));
                        var laneOrSlot = index % maxSimultaneous;
                        var displayStart = start.AddSeconds(index * EstimatedDisplayIntervalSeconds(overlay));
                        item["activeCountAtDisplay"] = Math.Min(index + 1, maxSimultaneous);
                        item["placementMode"] = "runtime";
                        item["selectedRegionIndex"] = null;
                        item["fallbackAnchor"] = null;
                        item["displayStartTime"] = displayStart.ToString("o");
                        if (overlay != null && overlay.DisplayMode == OverlayDisplayMode.Flow)
                        {
                            item["laneIndex"] = laneOrSlot;
                            item["removedReason"] = "animationCompleted";
                        }
                        else if (overlay != null && overlay.DisplayMode == OverlayDisplayMode.Bubble)
                        {
                            item["slotIndex"] = laneOrSlot;
                            item["removedReason"] = "fadeCompleted";
                        }

                        if (overlay != null)
                        {
                            item["displayEndTime"] = displayStart.AddSeconds(EstimatedCycleSeconds(overlay)).ToString("o");
                        }
                    }

                    return item;
                })
                .ToList();
        }

        private static double EstimatedCycleSeconds(OverlaySettings overlay)
        {
            if (overlay == null)
            {
                return 0;
            }

            if (overlay.DisplayMode == OverlayDisplayMode.Bubble)
            {
                var fadeIn = overlay.BubbleFadeInSeconds > 0 ? overlay.BubbleFadeInSeconds : overlay.BubbleFadeSeconds;
                var fadeOut = overlay.BubbleFadeOutSeconds > 0 ? overlay.BubbleFadeOutSeconds : overlay.BubbleFadeSeconds;
                return Math.Max(0, fadeIn) + Math.Max(0, overlay.BubbleLifetimeSeconds) + Math.Max(0, fadeOut);
            }

            if (overlay.DisplayMode == OverlayDisplayMode.Flow)
            {
                return 14;
            }

            return Math.Max(0, overlay.CommentLifetimeSeconds);
        }

        private static double EstimatedDisplayIntervalSeconds(OverlaySettings overlay)
        {
            if (overlay == null)
            {
                return 1.2;
            }

            var min = overlay.CommentDisplayIntervalMinSeconds > 0
                ? overlay.CommentDisplayIntervalMinSeconds
                : overlay.CommentDisplayIntervalSeconds;
            var max = overlay.CommentDisplayIntervalMaxSeconds > 0
                ? overlay.CommentDisplayIntervalMaxSeconds
                : overlay.CommentDisplayIntervalSeconds;
            min = Math.Max(0.1, min);
            max = Math.Max(0.1, max);
            return (min + max) / 2.0;
        }

        private static Dictionary<string, object> ToOverlaySettingsSnapshot(OverlaySettings overlay)
        {
            if (overlay == null)
            {
                return null;
            }

            return new Dictionary<string, object>
            {
                ["commentDisplayIntervalMinSeconds"] = overlay.CommentDisplayIntervalMinSeconds,
                ["commentDisplayIntervalMaxSeconds"] = overlay.CommentDisplayIntervalMaxSeconds,
                ["whitespacePlacementProbability"] = overlay.WhitespacePlacementProbability,
                ["commentFontSize"] = overlay.CommentFontSize,
                ["commentTextColor"] = overlay.CommentTextColor,
                ["commentBackgroundColor"] = overlay.CommentBackgroundColor,
                ["flowSpeedPixelsPerSecond"] = overlay.FlowSpeedPixelsPerSecond,
                ["bubbleFadeInSeconds"] = overlay.BubbleFadeInSeconds,
                ["bubbleFadeOutSeconds"] = overlay.BubbleFadeOutSeconds,
                ["presentationMonitorDeviceName"] = overlay.PresentationMonitorDeviceName,
                ["presentationMonitorIndex"] = overlay.PresentationMonitorIndex
            };
        }

        private static Dictionary<string, object> ToLogOverlayBounds(OverlayBounds? bounds)
        {
            if (!bounds.HasValue)
            {
                return null;
            }

            var value = bounds.Value;
            return new Dictionary<string, object>
            {
                ["source"] = value.Source,
                ["slideShowWindowHwnd"] = value.Hwnd,
                ["left"] = value.Left,
                ["top"] = value.Top,
                ["right"] = value.Right,
                ["bottom"] = value.Bottom,
                ["width"] = value.Width,
                ["height"] = value.Height,
                ["dpiX"] = value.DpiX,
                ["dpiY"] = value.DpiY
            };
        }

        private static IReadOnlyList<Dictionary<string, object>> ToLogScreens()
        {
            try
            {
                return System.Windows.Forms.Screen.AllScreens
                    .Select((screen, index) => new Dictionary<string, object>
                    {
                        ["index"] = index,
                        ["deviceName"] = screen.DeviceName,
                        ["primary"] = screen.Primary,
                        ["boundsLeft"] = screen.Bounds.Left,
                        ["boundsTop"] = screen.Bounds.Top,
                        ["boundsRight"] = screen.Bounds.Right,
                        ["boundsBottom"] = screen.Bounds.Bottom,
                        ["boundsWidth"] = screen.Bounds.Width,
                        ["boundsHeight"] = screen.Bounds.Height,
                        ["workingAreaLeft"] = screen.WorkingArea.Left,
                        ["workingAreaTop"] = screen.WorkingArea.Top,
                        ["workingAreaRight"] = screen.WorkingArea.Right,
                        ["workingAreaBottom"] = screen.WorkingArea.Bottom,
                        ["workingAreaWidth"] = screen.WorkingArea.Width,
                        ["workingAreaHeight"] = screen.WorkingArea.Height
                    })
                    .ToList();
            }
            catch
            {
                return new List<Dictionary<string, object>>();
            }
        }

        private static Dictionary<string, object> ToLogRegion(WhitespaceRegion region)
        {
            if (region == null)
            {
                return null;
            }

            return new Dictionary<string, object>
            {
                ["x"] = region.X,
                ["y"] = region.Y,
                ["width"] = region.Width,
                ["height"] = region.Height,
                ["score"] = region.Score
            };
        }

        private static Dictionary<string, object> ToPreloadData(
            SlideSnapshot snapshot,
            SlidePreloadResult result,
            bool cacheHit)
        {
            var cacheKey = result?.CacheKey ?? snapshot?.CacheKey;
            return new Dictionary<string, object>
            {
                ["cacheKeyHash"] = HashPath(cacheKey),
                ["usedApi"] = result != null ? (object)result.UsedApi : null,
                ["preloadSucceeded"] = result != null ? (object)result.Succeeded : null,
                ["preloadTotalTimeMs"] = result != null ? (object)result.TotalPreloadTimeMs : null,
                ["slideExportTimeMs"] = result != null ? (object)result.SlideExportTimeMs : snapshot?.SlideExportTimeMs,
                ["textExtractTimeMs"] = result != null ? (object)result.TextExtractTimeMs : snapshot?.TextExtractTimeMs,
                ["whitespaceAnalysisTimeMs"] = result != null ? (object)result.WhitespaceAnalysisTimeMs : null,
                ["commentGenerationTimeMs"] = result != null ? (object)result.CommentGenerationTimeMs : null,
                ["cacheHit"] = cacheHit,
                ["errorMessage"] = result?.ErrorMessage
            };
        }

        private static string HashPath(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            using (var sha256 = SHA256.Create())
            {
                var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(value));
                var builder = new StringBuilder(bytes.Length * 2);
                foreach (var b in bytes)
                {
                    builder.Append(b.ToString("x2"));
                }

                return builder.ToString();
            }
        }
    }
}
