using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using SlideAudienceAddIn.Models;
using SlideAudienceAddIn.Overlay;
using SlideAudienceAddIn.Ribbon;
using SlideAudienceAddIn.Services;

namespace SlideAudienceAddIn
{
    public partial class ThisAddIn
    {
        private OverlayController _overlayController;
        private SlideExporter _slideExporter;
        private SlideTextExtractor _slideTextExtractor;
        private SlideWhitespaceAnalyzer _slideWhitespaceAnalyzer;
        private CommentGenerationService _commentGenerationService;
        private SlidePreloadService _slidePreloadService;
        private ExperimentLogger _experimentLogger;
        private AppSettings _settings;
        private CancellationTokenSource _commentGenerationCancellation;
        private CancellationTokenSource _preloadCancellation;
        private bool _isEnabled = true;
        private string _lastLoggedDisplayMode;
        private string _conditionPreset = "None";

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Debug.WriteLine("[SlideAudience] Add-in startup");

            _settings = AppSettings.Load();
            _isEnabled = _settings.Enabled;
            _overlayController = new OverlayController();
            _overlayController.ApplySettings(_settings.Overlay);
            _slideExporter = new SlideExporter();
            _slideTextExtractor = new SlideTextExtractor();
            _slideWhitespaceAnalyzer = new SlideWhitespaceAnalyzer(_settings);
            _commentGenerationService = new CommentGenerationService(_settings);
            _slidePreloadService = new SlidePreloadService(_slideWhitespaceAnalyzer, _commentGenerationService);
            _experimentLogger = new ExperimentLogger(_settings);
            _preloadCancellation = new CancellationTokenSource();

            // PowerPoint slideshow events
            this.Application.SlideShowBegin += Application_SlideShowBegin;
            this.Application.SlideShowNextSlide += Application_SlideShowNextSlide;
            this.Application.SlideShowEnd += Application_SlideShowEnd;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Debug.WriteLine("[SlideAudience] Add-in shutdown");

            try
            {
                this.Application.SlideShowBegin -= Application_SlideShowBegin;
                this.Application.SlideShowNextSlide -= Application_SlideShowNextSlide;
                this.Application.SlideShowEnd -= Application_SlideShowEnd;

                _overlayController?.Close();
                _overlayController = null;
                _commentGenerationCancellation?.Cancel();
                _commentGenerationCancellation?.Dispose();
                _commentGenerationCancellation = null;
                _preloadCancellation?.Cancel();
                _preloadCancellation?.Dispose();
                _preloadCancellation = null;
                
            }
            catch
            {
                // Ignore shutdown cleanup errors.
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SlideAudienceRibbon(this);
        }

        public bool IsEnabled
        {
            get { return _isEnabled; }
            set
            {
                _isEnabled = value;
                if (_settings != null)
                {
                    _settings.Enabled = value;
                    SaveSettings("Enabled", value);
                }

                if (!value)
                {
                    _commentGenerationCancellation?.Cancel();
                    _overlayController?.StartNewSlideSession();
                    _overlayController?.Clear("disabled");
                    ClearPreloadCache("disabled");
                }

                Debug.WriteLine($"[SlideAudience] Enable changed: {value}");
                _experimentLogger?.LogSettingsChanged("Enabled", value);
            }
        }

        public AppSettings CurrentSettings
        {
            get { return _settings; }
        }

        public string ConditionPreset
        {
            get { return _conditionPreset; }
        }

        public void SetDisplayMode(OverlayDisplayMode displayMode)
        {
            var previous = _settings.Overlay.DisplayMode.ToString();
            _settings.Overlay.DisplayMode = displayMode;
            ApplyOverlaySettingsAndSave("Overlay.DisplayMode", displayMode.ToString());
            Debug.WriteLine($"[SlideAudience] DisplayMode changed: {displayMode}");
            _experimentLogger?.LogOverlayModeChanged(previous, displayMode.ToString());
        }

        public void SetConditionPreset(string preset)
        {
            _conditionPreset = string.IsNullOrWhiteSpace(preset) ? "None" : preset;
            Debug.WriteLine($"[SlideAudience] ConditionPreset changed: {_conditionPreset}");

            switch (_conditionPreset)
            {
                case "Panel":
                    _isEnabled = true;
                    _settings.Enabled = true;
                    _settings.Comments.MaxCommentsPerSlide = 10;
                    _settings.Overlay.DisplayMode = OverlayDisplayMode.Panel;
                    _settings.Overlay.MaxSimultaneousComments = 3;
                    _settings.Overlay.CommentLifetimeSeconds = 7;
                    break;
                case "Flow":
                    _isEnabled = true;
                    _settings.Enabled = true;
                    _settings.Comments.MaxCommentsPerSlide = 10;
                    _settings.Overlay.DisplayMode = OverlayDisplayMode.Flow;
                    _settings.Overlay.MaxSimultaneousComments = 3;
                    _settings.Overlay.FlowSpeedPixelsPerSecond = 80;
                    _settings.Overlay.CommentDisplayIntervalSeconds = 1.2;
                    _settings.Overlay.CommentDisplayIntervalMinSeconds = 1.0;
                    _settings.Overlay.CommentDisplayIntervalMaxSeconds = 2.5;
                    _settings.Overlay.UseWhitespaceAwarePlacement = true;
                    break;
                case "Bubble":
                    _isEnabled = true;
                    _settings.Enabled = true;
                    _settings.Comments.MaxCommentsPerSlide = 10;
                    _settings.Overlay.DisplayMode = OverlayDisplayMode.Bubble;
                    _settings.Overlay.MaxSimultaneousComments = 3;
                    _settings.Overlay.BubbleLifetimeSeconds = 9;
                    _settings.Overlay.BubbleFadeSeconds = 1.0;
                    _settings.Overlay.BubbleFadeInSeconds = 1.0;
                    _settings.Overlay.BubbleFadeOutSeconds = 1.0;
                    _settings.Overlay.CommentDisplayIntervalSeconds = 1.2;
                    _settings.Overlay.CommentDisplayIntervalMinSeconds = 1.0;
                    _settings.Overlay.CommentDisplayIntervalMaxSeconds = 2.5;
                    _settings.Overlay.UseWhitespaceAwarePlacement = true;
                    break;
                case "Debug":
                    _isEnabled = true;
                    _settings.Enabled = true;
                    _settings.Gemini.EnableApi = false;
                    _settings.Overlay.DisplayMode = OverlayDisplayMode.Panel;
                    _settings.Experiment.EnableLogging = true;
                    break;
                default:
                    _isEnabled = false;
                    _settings.Enabled = false;
                    _commentGenerationCancellation?.Cancel();
                    _overlayController?.StartNewSlideSession();
                    _overlayController?.Clear("conditionNone");
                    ClearPreloadCache("conditionNone");
                    break;
            }

            ApplyOverlaySettingsAndSave("ConditionPreset", _conditionPreset);
            _experimentLogger?.LogConditionPresetChanged(_conditionPreset);
        }

        public void SetEnableApi(bool enableApi)
        {
            _settings.Gemini.EnableApi = enableApi;
            Debug.WriteLine($"[SlideAudience] EnableApi changed: {enableApi}");
            ClearPreloadCache("EnableApiChanged");
            SaveSettings("Gemini.EnableApi", enableApi);
            _experimentLogger?.LogSettingsChanged("Gemini.EnableApi", enableApi);
        }

        public bool HasGeminiApiKey()
        {
            var variableName = _settings?.Gemini?.ApiKeyEnvironmentVariable;
            return !string.IsNullOrWhiteSpace(variableName) &&
                !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(variableName));
        }

        public void SetWhitespaceAwarePlacement(bool enabled)
        {
            _settings.Overlay.UseWhitespaceAwarePlacement = enabled;
            Debug.WriteLine($"[SlideAudience] UseWhitespaceAwarePlacement changed: {enabled}");
            ClearPreloadCache("WhitespaceAwarePlacementChanged");
            ApplyOverlaySettingsAndSave("Overlay.UseWhitespaceAwarePlacement", enabled);
        }

        public void SetMaxCommentsPerSlide(int count)
        {
            _settings.Comments.MaxCommentsPerSlide = Math.Max(1, count);
            Debug.WriteLine($"[SlideAudience] MaxCommentsPerSlide changed: {_settings.Comments.MaxCommentsPerSlide}");
            ClearPreloadCache("MaxCommentsPerSlideChanged");
            SaveSettings("Comments.MaxCommentsPerSlide", _settings.Comments.MaxCommentsPerSlide);
            _experimentLogger?.LogSettingsChanged("Comments.MaxCommentsPerSlide", _settings.Comments.MaxCommentsPerSlide);
        }

        public void SetMaxSimultaneousComments(int count)
        {
            _settings.Overlay.MaxSimultaneousComments = Math.Max(1, Math.Min(3, count));
            Debug.WriteLine($"[SlideAudience] MaxSimultaneousComments changed: {_settings.Overlay.MaxSimultaneousComments}");
            ApplyOverlaySettingsAndSave("Overlay.MaxSimultaneousComments", _settings.Overlay.MaxSimultaneousComments);
        }

        public void SetCommentDisplayIntervalMinSeconds(double seconds)
        {
            _settings.Overlay.CommentDisplayIntervalMinSeconds = Clamp(seconds, 0.1, 30);
            if (_settings.Overlay.CommentDisplayIntervalMinSeconds > _settings.Overlay.CommentDisplayIntervalMaxSeconds)
            {
                _settings.Overlay.CommentDisplayIntervalMaxSeconds = _settings.Overlay.CommentDisplayIntervalMinSeconds;
                Debug.WriteLine("[SlideAudience] CommentDisplayIntervalMaxSeconds aligned to Min because min > max");
            }

            Debug.WriteLine($"[SlideAudience] CommentDisplayIntervalMinSeconds changed: {_settings.Overlay.CommentDisplayIntervalMinSeconds}");
            ApplyOverlaySettingsAndSave("Overlay.CommentDisplayIntervalMinSeconds", _settings.Overlay.CommentDisplayIntervalMinSeconds);
        }

        public void SetCommentDisplayIntervalMaxSeconds(double seconds)
        {
            _settings.Overlay.CommentDisplayIntervalMaxSeconds = Clamp(seconds, 0.1, 30);
            if (_settings.Overlay.CommentDisplayIntervalMinSeconds > _settings.Overlay.CommentDisplayIntervalMaxSeconds)
            {
                _settings.Overlay.CommentDisplayIntervalMinSeconds = _settings.Overlay.CommentDisplayIntervalMaxSeconds;
                Debug.WriteLine("[SlideAudience] CommentDisplayIntervalMinSeconds aligned to Max because min > max");
            }

            Debug.WriteLine($"[SlideAudience] CommentDisplayIntervalMaxSeconds changed: {_settings.Overlay.CommentDisplayIntervalMaxSeconds}");
            ApplyOverlaySettingsAndSave("Overlay.CommentDisplayIntervalMaxSeconds", _settings.Overlay.CommentDisplayIntervalMaxSeconds);
        }

        public void SetWhitespacePlacementProbability(double probability)
        {
            _settings.Overlay.WhitespacePlacementProbability = Clamp(probability, 0, 1);
            Debug.WriteLine($"[SlideAudience] WhitespacePlacementProbability changed: {_settings.Overlay.WhitespacePlacementProbability}");
            ClearPreloadCache("WhitespacePlacementProbabilityChanged");
            ApplyOverlaySettingsAndSave("Overlay.WhitespacePlacementProbability", _settings.Overlay.WhitespacePlacementProbability);
        }

        public void SetCommentFontSize(double fontSize)
        {
            _settings.Overlay.CommentFontSize = Clamp(fontSize, 16, 48);
            Debug.WriteLine($"[SlideAudience] CommentFontSize changed: {_settings.Overlay.CommentFontSize}");
            ApplyOverlaySettingsAndSave("Overlay.CommentFontSize", _settings.Overlay.CommentFontSize);
        }

        public void SetCommentTextColor(string color)
        {
            _settings.Overlay.CommentTextColor = string.IsNullOrWhiteSpace(color) ? "#FFFFFFFF" : color;
            Debug.WriteLine($"[SlideAudience] CommentTextColor changed: {_settings.Overlay.CommentTextColor}");
            ApplyOverlaySettingsAndSave("Overlay.CommentTextColor", _settings.Overlay.CommentTextColor);
        }

        public void SetCommentBackgroundColor(string color)
        {
            _settings.Overlay.CommentBackgroundColor = string.IsNullOrWhiteSpace(color) ? "#99000000" : color;
            Debug.WriteLine($"[SlideAudience] CommentBackgroundColor changed: {_settings.Overlay.CommentBackgroundColor}");
            ApplyOverlaySettingsAndSave("Overlay.CommentBackgroundColor", _settings.Overlay.CommentBackgroundColor);
        }

        public void SetFlowSpeedPixelsPerSecond(double speed)
        {
            _settings.Overlay.FlowSpeedPixelsPerSecond = Clamp(speed, 30, 200);
            Debug.WriteLine($"[SlideAudience] FlowSpeedPixelsPerSecond changed: {_settings.Overlay.FlowSpeedPixelsPerSecond}");
            ApplyOverlaySettingsAndSave("Overlay.FlowSpeedPixelsPerSecond", _settings.Overlay.FlowSpeedPixelsPerSecond);
        }

        public void SetBubbleFadeInSeconds(double seconds)
        {
            _settings.Overlay.BubbleFadeInSeconds = Clamp(seconds, 0.1, 10);
            Debug.WriteLine($"[SlideAudience] BubbleFadeInSeconds changed: {_settings.Overlay.BubbleFadeInSeconds}");
            ApplyOverlaySettingsAndSave("Overlay.BubbleFadeInSeconds", _settings.Overlay.BubbleFadeInSeconds);
        }

        public void SetBubbleFadeOutSeconds(double seconds)
        {
            _settings.Overlay.BubbleFadeOutSeconds = Clamp(seconds, 0.1, 10);
            Debug.WriteLine($"[SlideAudience] BubbleFadeOutSeconds changed: {_settings.Overlay.BubbleFadeOutSeconds}");
            ApplyOverlaySettingsAndSave("Overlay.BubbleFadeOutSeconds", _settings.Overlay.BubbleFadeOutSeconds);
        }

        public void SetPresentationMonitor(string deviceName, int screenIndex)
        {
            _settings.Overlay.PresentationMonitorDeviceName = string.IsNullOrWhiteSpace(deviceName)
                ? string.Empty
                : deviceName;
            _settings.Overlay.PresentationMonitorIndex = string.IsNullOrWhiteSpace(deviceName)
                ? -1
                : screenIndex;

            Debug.WriteLine($"[SlideAudience] Presentation monitor changed: device={_settings.Overlay.PresentationMonitorDeviceName}, index={_settings.Overlay.PresentationMonitorIndex}");
            ApplyOverlaySettingsAndSave("Overlay.PresentationMonitor", $"{_settings.Overlay.PresentationMonitorIndex}:{_settings.Overlay.PresentationMonitorDeviceName}");
        }

        public void ClearOverlay()
        {
            Debug.WriteLine("[SlideAudience] ClearOverlay clicked");
            _commentGenerationCancellation?.Cancel();
            _overlayController?.StartNewSlideSession();
            _overlayController?.Clear("manual");
            _experimentLogger?.LogCommentsHidden();
        }

        public void SetPreloadNextSlide(bool enabled)
        {
            _settings.Preload.EnablePreloadNextSlide = enabled;
            Debug.WriteLine($"[SlideAudience] PreloadNextSlide changed: {enabled}");
            if (!enabled)
            {
                ClearPreloadCache("PreloadDisabled");
            }

            SaveSettings("Preload.EnablePreloadNextSlide", enabled);
            _experimentLogger?.LogSettingsChanged("Preload.EnablePreloadNextSlide", enabled);
        }

        public void SetPreloadSlideCount(int count)
        {
            _settings.Preload.PreloadSlideCount = Math.Max(1, Math.Min(3, count));
            Debug.WriteLine($"[SlideAudience] PreloadSlideCount changed: {_settings.Preload.PreloadSlideCount}");
            SaveSettings("Preload.PreloadSlideCount", _settings.Preload.PreloadSlideCount);
            _experimentLogger?.LogSettingsChanged("Preload.PreloadSlideCount", _settings.Preload.PreloadSlideCount);
        }

        public void ClearPreloadCache()
        {
            ClearPreloadCache("manual");
        }

        public void OpenLogFolder()
        {
            Debug.WriteLine("[SlideAudience] OpenLogFolder clicked");
            var logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "SlideAudience",
                "logs");
            Directory.CreateDirectory(logDir);
            Process.Start("explorer.exe", logDir);
        }

        public void TestCurrentSlide()
        {
            Debug.WriteLine("[SlideAudience] Test Current Slide clicked");
            try
            {
                if (Application.SlideShowWindows.Count > 0)
                {
                    ShowOverlayForCurrentSlide(Application.SlideShowWindows[1], "test");
                    return;
                }

                Debug.WriteLine("[SlideAudience] Test Current Slide skipped: slideshow is not running");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SlideAudience] Test Current Slide failed");
                Debug.WriteLine(ex.ToString());
                _experimentLogger?.LogError(null, null, ex);
            }
        }

        public void GenerateForCurrentSlide()
        {
            TestCurrentSlide();
        }

        public string GetSettingsSummary()
        {
            return $"Enabled={_isEnabled}, DisplayMode={_settings?.Overlay?.DisplayMode}, EnableApi={_settings?.Gemini?.EnableApi}";
        }

        private void ApplyOverlaySettingsAndSave(string settingName, object value)
        {
            _overlayController?.ApplySettings(_settings.Overlay);
            SaveSettings(settingName, value);
            _experimentLogger?.LogSettingsChanged(settingName, value);
        }

        private void SaveSettings(string settingName, object value)
        {
            try
            {
                _settings.SaveLocal();
                Debug.WriteLine("[SlideAudience] settings saved to appsettings.local.json");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] failed to save settings after {settingName}={value}");
                Debug.WriteLine(ex.ToString());
                _experimentLogger?.LogError(null, null, ex);
            }
        }

        private static double Clamp(double value, double min, double max)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return min;
            }

            return Math.Max(min, Math.Min(max, value));
        }

        private static string FormatOverlayBounds(OverlayBounds? bounds)
        {
            if (!bounds.HasValue)
            {
                return "(null)";
            }

            var value = bounds.Value;
            return $"Source={value.Source}, Hwnd=0x{value.Hwnd:X}, Left={value.Left:F0}, Top={value.Top:F0}, Right={value.Right:F0}, Bottom={value.Bottom:F0}, Width={value.Width:F0}, Height={value.Height:F0}";
        }

        private static void LogSlideShowEventWindow(string eventName, PowerPoint.SlideShowWindow window)
        {
            try
            {
                var hwnd = window != null ? window.HWND : 0;
                int? slideId = null;
                int? slideIndex = null;
                try
                {
                    var slide = window?.View?.Slide;
                    slideId = slide?.SlideID;
                    slideIndex = slide?.SlideIndex;
                }
                catch
                {
                    // Best-effort diagnostics only.
                }

                Debug.WriteLine($"[SlideAudience] {eventName} fired SlideShowWindow.HWND=0x{hwnd:X}, SlideID={slideId}, SlideIndex={slideIndex}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] {eventName} fired but window diagnostics failed");
                Debug.WriteLine(ex.ToString());
            }
        }

        private void Application_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            Debug.WriteLine("[SlideAudience] Slide show begin");
            LogSlideShowEventWindow("SlideShowBegin", Wn);
            _experimentLogger?.StartSession(GetPresentationPath(Wn));
            _lastLoggedDisplayMode = null;
            LogOverlayModeIfChanged();
            ShowOverlayForCurrentSlide(Wn, "begin");
        }

        private void Application_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            Debug.WriteLine("[SlideAudience] Slide show next slide");
            Debug.WriteLine("[SlideAudience] SlideShowNextSlide fired");
            LogSlideShowEventWindow("SlideShowNextSlide", Wn);
            LogOverlayModeIfChanged();
            ShowOverlayForCurrentSlide(Wn, "next");
        }

        private void Application_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            Debug.WriteLine("[SlideAudience] Slide show end");

            try
            {
                _commentGenerationCancellation?.Cancel();
                ClearPreloadCache("slideshowEnded");
                _experimentLogger?.LogCommentsHidden();
                _overlayController?.Close();

                string name = Pres != null ? Pres.Name : "(unknown presentation)";
                Debug.WriteLine($"[SlideAudience] Presentation ended: {name}");
                _experimentLogger?.EndSession();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] Error on slideshow end: {ex.Message}");
                _experimentLogger?.LogError(null, null, ex);
            }
        }

        private void ShowOverlayForCurrentSlide(PowerPoint.SlideShowWindow Wn, string eventName)
        {
            try
            {
                if (!_isEnabled)
                {
                    Debug.WriteLine($"[SlideAudience] {eventName}: add-in disabled");
                    return;
                }

                if (Wn == null || Wn.View == null || Wn.View.Slide == null)
                {
                    Debug.WriteLine($"[SlideAudience] {eventName}: slide window/view/slide is null");
                    return;
                }

                PowerPoint.Slide slide = Wn.View.Slide;

                int slideIndex = slide.SlideIndex;
                int slideId = slide.SlideID;
                string slideName = slide.Name;
                string exportedPath = null;
                string slideText = string.Empty;
                bool imageExportSucceeded = false;
                bool textExtractionSucceeded = false;
                long? slideExportTimeMs = null;
                long? whitespaceAnalysisTimeMs = null;
                SlideWhitespaceAnalysisResult whitespaceAnalysis = null;
                OverlayBounds? overlayBounds = _overlayController?.CaptureSlideShowBounds(Wn);

                Debug.WriteLine(
                    $"[SlideAudience] {eventName}: show overlay index={slideIndex}, id={slideId}, name={slideName}"
                );
                Debug.WriteLine($"[SlideAudience] SlideID/SlideIndex event={eventName}, SlideID={slideId}, SlideIndex={slideIndex}");
                Debug.WriteLine($"[SlideAudience] Overlay bounds captured event={eventName}, bounds={FormatOverlayBounds(overlayBounds)}");
                _experimentLogger?.LogOverlayDiagnostics(slideId, slideIndex, $"{eventName}:bounds_captured", overlayBounds);
                _experimentLogger?.LogSlideChanged(slideId, slideIndex);
                var slideSessionId = _overlayController?.StartNewSlideSession() ?? 0;
                _commentGenerationCancellation?.Cancel();
                _overlayController?.ResetForNewSlide("slideChanged");
                var presentationHash = HashValue(GetPresentationPath(Wn));
                var cacheKey = BuildCacheKey(presentationHash, slideId, slideIndex);

                if (_settings.Preload.EnablePreloadNextSlide &&
                    _slidePreloadService != null &&
                    _slidePreloadService.TryGetCached(cacheKey, out var cachedPreload) &&
                    cachedPreload != null &&
                    cachedPreload.Succeeded &&
                    cachedPreload.CommentResult?.Comments?.Count > 0)
                {
                    Debug.WriteLine($"[SlideAudience] Using preloaded comments for slide {slideIndex}");
                    _experimentLogger?.LogPreloadUsed(cachedPreload, cacheHit: true);

                    if (_overlayController != null && !_overlayController.IsCurrentSession(slideSessionId))
                    {
                        Debug.WriteLine($"[SlideAudience] Ignored cached result because session changed: slideIndex={slideIndex}, sessionId={slideSessionId}");
                        return;
                    }

                    var cachedComments = cachedPreload.CommentResult.Comments
                        .Where(comment => comment != null && !string.IsNullOrWhiteSpace(comment.Text))
                        .ToList();
                    var cachedRegions = cachedPreload.WhitespaceAnalysis != null && cachedPreload.WhitespaceAnalysis.Succeeded
                        ? cachedPreload.WhitespaceAnalysis.Regions
                        : new List<WhitespaceRegion>();

                    Debug.WriteLine($"[SlideAudience] Comment generation success=True source=preload, slideId={slideId}, slideIndex={slideIndex}, count={cachedComments.Count}");
                    Debug.WriteLine($"[SlideAudience] Overlay display update requested source=preload, slideId={slideId}, slideIndex={slideIndex}, bounds={FormatOverlayBounds(overlayBounds)}");
                    _experimentLogger?.LogOverlayDiagnostics(slideId, slideIndex, "overlay_update_requested_preload", overlayBounds);
                    _overlayController?.ShowComments(overlayBounds, cachedComments, cachedRegions, slideSessionId);
                    _experimentLogger?.LogCommentsShown(slideId, slideIndex, cachedComments, cachedRegions);
                    StartPreloadForNextSlide(Wn, slideIndex, presentationHash);
                    return;
                }

                Debug.WriteLine($"[SlideAudience] No preload cache for slide {slideIndex}, generating now");
                _experimentLogger?.LogPreloadCacheMiss(slideId, slideIndex, cacheKey);

                try
                {
                    var exportStopwatch = Stopwatch.StartNew();
                    exportedPath = _slideExporter.ExportSlideAsPng(slide);
                    exportStopwatch.Stop();
                    slideExportTimeMs = exportStopwatch.ElapsedMilliseconds;
                    imageExportSucceeded = true;
                    Debug.WriteLine($"[SlideAudience] exported slide png: {exportedPath}");
                    Debug.WriteLine($"[SlideAudience] Slide.Export success=True slideId={slideId}, slideIndex={slideIndex}, path={exportedPath}");
                    Debug.WriteLine($"[SlideAudience] Slide export took {slideExportTimeMs} ms");
                }
                catch (Exception exportEx)
                {
                    Debug.WriteLine($"[SlideAudience] Slide.Export success=False slideId={slideId}, slideIndex={slideIndex}");
                    Debug.WriteLine($"[SlideAudience] failed to export slide png: {exportEx}");
                    _experimentLogger?.LogError(slideId, slideIndex, exportEx);
                }

                if (imageExportSucceeded)
                {
                    var whitespaceStopwatch = Stopwatch.StartNew();
                    whitespaceAnalysis = _slideWhitespaceAnalyzer?.Analyze(exportedPath);
                    whitespaceStopwatch.Stop();
                    whitespaceAnalysisTimeMs = whitespaceStopwatch.ElapsedMilliseconds;
                    Debug.WriteLine($"[SlideAudience] Whitespace analysis took {whitespaceAnalysisTimeMs} ms");
                }
                else
                {
                    whitespaceAnalysis = new SlideWhitespaceAnalysisResult
                    {
                        Succeeded = false,
                        FallbackReason = "slide image export failed",
                        Regions = new List<WhitespaceRegion>()
                    };
                    Debug.WriteLine("[SlideAudience] placement fallback reason=slide image export failed");
                }

                try
                {
                    slideText = _slideTextExtractor.ExtractText(slide);
                    textExtractionSucceeded = true;
                    Debug.WriteLine("[SlideAudience] extracted slide text:");
                    Debug.WriteLine(slideText);
                }
                catch (Exception textEx)
                {
                    Debug.WriteLine($"[SlideAudience] failed to extract slide text: {textEx}");
                    _experimentLogger?.LogError(slideId, slideIndex, textEx);
                }

                _experimentLogger?.LogSlideAnalyzed(
                    slideId,
                    slideIndex,
                    exportedPath,
                    slideText,
                    imageExportSucceeded,
                    textExtractionSucceeded,
                    whitespaceAnalysis,
                    slideExportTimeMs,
                    whitespaceAnalysisTimeMs);

                if (_overlayController != null && !_overlayController.IsCurrentSession(slideSessionId))
                {
                    Debug.WriteLine($"[SlideAudience] Ignored stale slide analysis before generation: slideIndex={slideIndex}, sessionId={slideSessionId}");
                    return;
                }

                ShowGeneratedComments(overlayBounds, slideId, slideIndex, exportedPath, slideText, whitespaceAnalysis, slideSessionId);
                StartPreloadForNextSlide(Wn, slideIndex, presentationHash);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] Error while showing overlay: {ex}");
                _experimentLogger?.LogError(null, null, ex);
            }
        }

        private async void ShowGeneratedComments(
            OverlayBounds? overlayBounds,
            int slideId,
            int slideIndex,
            string exportedPath,
            string slideText,
            SlideWhitespaceAnalysisResult whitespaceAnalysis,
            int slideSessionId)
        {
            try
            {
                Debug.WriteLine("[SlideAudience] ShowGeneratedComments called");
                Debug.WriteLine($"[SlideAudience] ShowGeneratedComments slideId={slideId}, slideIndex={slideIndex}");
                Debug.WriteLine($"[SlideAudience] ShowGeneratedComments exportedPath is null: {exportedPath == null}");
                Debug.WriteLine($"[SlideAudience] ShowGeneratedComments slideText length: {(slideText ?? string.Empty).Length}");

                var previousCancellation = _commentGenerationCancellation;
                var currentCancellation = new CancellationTokenSource();
                _commentGenerationCancellation = currentCancellation;

                previousCancellation?.Cancel();
                previousCancellation?.Dispose();

                Debug.WriteLine("[SlideAudience] GenerateForSlideAsync start");
                var generationStopwatch = Stopwatch.StartNew();
                var result = await _commentGenerationService.GenerateForSlideAsync(
                    slideId,
                    slideIndex,
                    exportedPath,
                    slideText,
                    currentCancellation.Token);
                generationStopwatch.Stop();
                Debug.WriteLine("[SlideAudience] GenerateForSlideAsync completed");
                Debug.WriteLine($"[SlideAudience] Comment generation took {generationStopwatch.ElapsedMilliseconds} ms");
                _experimentLogger?.LogCommentsGenerated(slideId, slideIndex, result);

                if (currentCancellation.IsCancellationRequested)
                {
                    Debug.WriteLine($"[SlideAudience] ShowGeneratedComments canceled after generation: slideIndex={slideIndex}");
                    return;
                }

                if (_overlayController != null && !_overlayController.IsCurrentSession(slideSessionId))
                {
                    Debug.WriteLine($"[SlideAudience] Ignored stale comments from previous slide: slideIndex={slideIndex}, sessionId={slideSessionId}");
                    return;
                }

                Debug.WriteLine($"[SlideAudience] GenerateForSlideAsync result is null: {result == null}");
                Debug.WriteLine($"[SlideAudience] GenerateForSlideAsync result.Comments count: {result?.Comments?.Count ?? 0}");
                Debug.WriteLine($"[SlideAudience] Comment generation success=True slideId={slideId}, slideIndex={slideIndex}, count={result?.Comments?.Count ?? 0}, usedApi={result?.UsedApi}, latencyMs={result?.LatencyMs}");

                var shownComments = result?.Comments?
                    .Where(comment => comment != null)
                    .Where(comment => !string.IsNullOrWhiteSpace(comment.Text))
                    .ToList() ?? new List<AudienceComment>();

                if (shownComments.Count == 0)
                {
                    Debug.WriteLine("[SlideAudience] generated comments were empty; using fallback comments");
                    shownComments = new List<AudienceComment>
                    {
                        AudienceComment.Create("fallback", "生成結果が空です", persona: "system"),
                        AudienceComment.Create("fallback", "解析処理は動作中", persona: "system"),
                        AudienceComment.Create("fallback", "次は生成確認です", persona: "system")
                    };
                }

                foreach (var comment in shownComments)
                {
                    Debug.WriteLine($"[SlideAudience] display comment.Text: {comment.Text}, persona={comment.Persona}");
                }

                Debug.WriteLine(
                    $"[SlideAudience] generated comments: count={shownComments.Count}, usedApi={result?.UsedApi}, latencyMs={result?.LatencyMs}"
                );

                if (!string.IsNullOrWhiteSpace(result?.ErrorMessage))
                {
                    Debug.WriteLine($"[SlideAudience] comment generation fallback: {result.ErrorMessage}");
                }

                Debug.WriteLine("[SlideAudience] calling OverlayController.ShowComments");
                var whitespaceRegions = whitespaceAnalysis != null && whitespaceAnalysis.Succeeded
                    ? whitespaceAnalysis.Regions
                    : new List<WhitespaceRegion>();
                var showStopwatch = Stopwatch.StartNew();
                Debug.WriteLine($"[SlideAudience] comments show start time={DateTimeOffset.UtcNow:o}");
                Debug.WriteLine($"[SlideAudience] Overlay display update requested slideId={slideId}, slideIndex={slideIndex}, bounds={FormatOverlayBounds(overlayBounds)}");
                _experimentLogger?.LogOverlayDiagnostics(slideId, slideIndex, "overlay_update_requested", overlayBounds);
                _overlayController?.ShowComments(overlayBounds, shownComments, whitespaceRegions, slideSessionId);
                showStopwatch.Stop();
                Debug.WriteLine($"[SlideAudience] ShowComments dispatch took {showStopwatch.ElapsedMilliseconds} ms");
                _experimentLogger?.LogCommentsShown(slideId, slideIndex, shownComments, whitespaceRegions);
                Debug.WriteLine("[SlideAudience] OverlayController.ShowComments returned");
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine($"[SlideAudience] comment generation canceled: slideIndex={slideIndex}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] Comment generation success=False slideId={slideId}, slideIndex={slideIndex}");
                Debug.WriteLine($"[SlideAudience] failed to generate comments: {ex}");
                Debug.WriteLine(ex.ToString());
                _experimentLogger?.LogError(slideId, slideIndex, ex);
                if (_overlayController != null && !_overlayController.IsCurrentSession(slideSessionId))
                {
                    Debug.WriteLine($"[SlideAudience] Ignored stale fallback overlay after generation error: slideIndex={slideIndex}, sessionId={slideSessionId}");
                    return;
                }

                _overlayController?.Show(overlayBounds, slideIndex);
            }
        }

        private void StartPreloadForNextSlide(
            PowerPoint.SlideShowWindow slideShowWindow,
            int currentSlideIndex,
            string presentationHash)
        {
            if (!_isEnabled || _settings?.Preload?.EnablePreloadNextSlide != true)
            {
                Debug.WriteLine("[SlideAudience] Preload skipped: disabled");
                return;
            }

            if (_slidePreloadService == null)
            {
                Debug.WriteLine("[SlideAudience] Preload skipped: service missing");
                return;
            }

            try
            {
                var nextSlideIndex = currentSlideIndex + 1;
                var presentation = slideShowWindow?.Presentation;
                if (presentation == null || nextSlideIndex > presentation.Slides.Count)
                {
                    Debug.WriteLine($"[SlideAudience] Preload skipped: no next slide after index={currentSlideIndex}");
                    return;
                }

                var nextSlide = presentation.Slides[nextSlideIndex];
                var snapshot = CreateSlideSnapshot(nextSlide, nextSlideIndex, presentationHash);
                _experimentLogger?.LogPreloadStarted(snapshot);

                var cancellationToken = _preloadCancellation?.Token ?? CancellationToken.None;
                var preloadTask = _slidePreloadService.PreloadSnapshotAsync(snapshot, _settings, cancellationToken);
                preloadTask.ContinueWith(task =>
                {
                    if (task.IsCanceled)
                    {
                        Debug.WriteLine($"[SlideAudience] Preload cancelled for slide {snapshot.SlideIndex}");
                        _experimentLogger?.LogPreloadCancelled(snapshot);
                        return;
                    }

                    if (task.IsFaulted)
                    {
                        Debug.WriteLine($"[SlideAudience] Preload faulted for slide {snapshot.SlideIndex}: {task.Exception}");
                        _experimentLogger?.LogPreloadFailed(snapshot, task.Exception);
                        return;
                    }

                    var result = task.Result;
                    if (result != null && result.Succeeded)
                    {
                        Debug.WriteLine($"[SlideAudience] Preload completed for slide {result.SlideIndex} in {result.TotalPreloadTimeMs} ms");
                        _experimentLogger?.LogPreloadCompleted(result);
                    }
                    else
                    {
                        Debug.WriteLine($"[SlideAudience] Preload failed for slide {snapshot.SlideIndex}: {result?.ErrorMessage}");
                        _experimentLogger?.LogPreloadFailed(result ?? new SlidePreloadResult
                        {
                            CacheKey = snapshot.CacheKey,
                            SlideIndex = snapshot.SlideIndex,
                            SlideId = snapshot.SlideId,
                            ErrorMessage = "preload result missing",
                            Succeeded = false,
                            CreatedAtUtc = DateTime.UtcNow
                        });
                    }
                }, TaskScheduler.Default);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] Failed to start preload for next slide: {ex}");
                Debug.WriteLine(ex.ToString());
                _experimentLogger?.LogError(null, null, ex);
            }
        }

        private SlideSnapshot CreateSlideSnapshot(
            PowerPoint.Slide slide,
            int slideIndex,
            string presentationHash)
        {
            if (slide == null)
            {
                throw new ArgumentNullException(nameof(slide));
            }

            var slideId = slide.SlideID;
            string imagePath = null;
            var slideText = string.Empty;
            long exportTimeMs = 0;
            long textExtractTimeMs = 0;

            var exportStopwatch = Stopwatch.StartNew();
            imagePath = _slideExporter.ExportSlideAsPng(slide);
            exportStopwatch.Stop();
            exportTimeMs = exportStopwatch.ElapsedMilliseconds;
            Debug.WriteLine($"[SlideAudience] Preload snapshot export took {exportTimeMs} ms, slideIndex={slideIndex}");

            try
            {
                var textStopwatch = Stopwatch.StartNew();
                slideText = _slideTextExtractor.ExtractText(slide);
                textStopwatch.Stop();
                textExtractTimeMs = textStopwatch.ElapsedMilliseconds;
                Debug.WriteLine($"[SlideAudience] Preload snapshot text extract took {textExtractTimeMs} ms, slideIndex={slideIndex}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] Preload snapshot text extract failed slideIndex={slideIndex}: {ex.Message}");
                Debug.WriteLine(ex.ToString());
                textExtractTimeMs = 0;
                slideText = string.Empty;
            }

            return new SlideSnapshot
            {
                CacheKey = BuildCacheKey(presentationHash, slideId, slideIndex),
                PresentationHash = presentationHash,
                SlideIndex = slideIndex,
                SlideId = slideId,
                ImagePath = imagePath,
                SlideText = slideText,
                SlideExportTimeMs = exportTimeMs,
                TextExtractTimeMs = textExtractTimeMs,
                CreatedAtUtc = DateTime.UtcNow
            };
        }

        private string BuildCacheKey(string presentationHash, int slideId, int slideIndex)
        {
            var settingsHash = HashValue(string.Join("|", new[]
            {
                _settings.Comments.Mode,
                _settings.Comments.MaxCommentsPerSlide.ToString(),
                _settings.Comments.MinCharacters.ToString(),
                _settings.Comments.MaxCharacters.ToString(),
                _settings.Comments.Language,
                _settings.Gemini.EnableApi.ToString(),
                _settings.Gemini.Model,
                _settings.Overlay.WhitespaceGridColumns.ToString(),
                _settings.Overlay.WhitespaceGridRows.ToString(),
                _settings.Overlay.UseWhitespaceAwarePlacement.ToString(),
                _settings.Overlay.WhitespacePlacementProbability.ToString("F3")
            }));

            return string.Join(":", new[]
            {
                string.IsNullOrWhiteSpace(presentationHash) ? "unknown" : presentationHash,
                slideId.ToString(),
                slideIndex.ToString(),
                settingsHash
            });
        }

        private void ClearPreloadCache(string reason)
        {
            try
            {
                _preloadCancellation?.Cancel();
                _preloadCancellation?.Dispose();
                _preloadCancellation = new CancellationTokenSource();
                _slidePreloadService?.Clear();
                Debug.WriteLine($"[SlideAudience] Preload cache cleared reason={reason}");
                _experimentLogger?.LogPreloadCacheCleared(reason);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] ClearPreloadCache failed reason={reason}");
                Debug.WriteLine(ex.ToString());
                _experimentLogger?.LogError(null, null, ex);
            }
        }

        private void LogOverlayModeIfChanged()
        {
            var currentMode = _settings?.Overlay?.DisplayMode.ToString();
            if (string.IsNullOrWhiteSpace(currentMode))
            {
                return;
            }

            if (!string.Equals(_lastLoggedDisplayMode, currentMode, StringComparison.OrdinalIgnoreCase))
            {
                _experimentLogger?.LogOverlayModeChanged(_lastLoggedDisplayMode, currentMode);
                _lastLoggedDisplayMode = currentMode;
            }
        }

        private static string GetPresentationPath(PowerPoint.SlideShowWindow window)
        {
            try
            {
                return window?.Presentation?.FullName;
            }
            catch
            {
                return null;
            }
        }

        private static string HashValue(string value)
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

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
