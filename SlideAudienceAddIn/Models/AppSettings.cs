using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace SlideAudienceAddIn.Models
{
    public class AppSettings
    {
        public GeminiSettings Gemini { get; set; } = new GeminiSettings();

        public CommentSettings Comments { get; set; } = new CommentSettings();

        public OverlaySettings Overlay { get; set; } = new OverlaySettings();

        public ExperimentSettings Experiment { get; set; } = new ExperimentSettings();

        public PreloadSettings Preload { get; set; } = new PreloadSettings();

        public bool Enabled { get; set; } = true;

        public static AppSettings Load()
        {
            var settings = new AppSettings();
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var configDir = Path.Combine(baseDir, "Config");
            var examplePath = Path.Combine(configDir, "appsettings.example.json");
            var localPath = Path.Combine(configDir, "appsettings.local.json");
            var path = File.Exists(localPath) ? localPath : examplePath;
            Debug.WriteLine($"[SlideAudience] AppSettings.Load config path: {path}");

            if (!File.Exists(path))
            {
                Debug.WriteLine("[SlideAudience] AppSettings.Load config file not found; using default settings");
                settings.ApplyEnvironmentOverrides();
                return settings;
            }

            try
            {
                var serializer = new JavaScriptSerializer();
                var root = serializer.Deserialize<Dictionary<string, object>>(File.ReadAllText(path));

                settings.ReadGemini(root);
                settings.ReadComments(root);
                settings.ReadOverlay(root);
                settings.ReadExperiment(root);
                settings.ReadPreload(root);
                settings.ApplyEnvironmentOverrides();
                Debug.WriteLine($"[SlideAudience] AppSettings.Load Gemini.EnableApi={settings.Gemini.EnableApi}, Gemini.Model={settings.Gemini.Model}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SlideAudience] AppSettings.Load failed; using default settings");
                Debug.WriteLine(ex.ToString());
                settings = new AppSettings();
                settings.ApplyEnvironmentOverrides();
            }

            return settings;
        }

        public void SaveLocal()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var configDir = Path.Combine(baseDir, "Config");
            Directory.CreateDirectory(configDir);
            SaveTo(Path.Combine(configDir, "appsettings.local.json"));
        }

        public void SaveTo(string path)
        {
            var root = new Dictionary<string, object>
            {
                ["Gemini"] = new Dictionary<string, object>
                {
                    ["Model"] = Gemini.Model,
                    ["ApiKeyEnvironmentVariable"] = Gemini.ApiKeyEnvironmentVariable,
                    ["TimeoutSeconds"] = Gemini.TimeoutSeconds,
                    ["EnableApi"] = Gemini.EnableApi
                },
                ["Comments"] = new Dictionary<string, object>
                {
                    ["Mode"] = Comments.Mode,
                    ["MaxCommentsPerSlide"] = Comments.MaxCommentsPerSlide,
                    ["MinCharacters"] = Comments.MinCharacters,
                    ["MaxCharacters"] = Comments.MaxCharacters,
                    ["Language"] = Comments.Language
                },
                ["Overlay"] = new Dictionary<string, object>
                {
                    ["Position"] = Overlay.Position,
                    ["Width"] = Overlay.Width,
                    ["MarginRight"] = Overlay.MarginRight,
                    ["MarginBottom"] = Overlay.MarginBottom,
                    ["FontSize"] = Overlay.FontSize,
                    ["CommentFontSize"] = Overlay.CommentFontSize,
                    ["CommentTextColor"] = Overlay.CommentTextColor,
                    ["CommentBackgroundColor"] = Overlay.CommentBackgroundColor,
                    ["UseNiconicoFlow"] = Overlay.UseNiconicoFlow,
                    ["DisplayMode"] = Overlay.DisplayMode.ToString(),
                    ["CommentLifetimeSeconds"] = Overlay.CommentLifetimeSeconds,
                    ["FlowSpeedPixelsPerSecond"] = Overlay.FlowSpeedPixelsPerSecond,
                    ["BubbleLifetimeSeconds"] = Overlay.BubbleLifetimeSeconds,
                    ["BubbleFadeSeconds"] = Overlay.BubbleFadeSeconds,
                    ["BubbleFadeInSeconds"] = Overlay.BubbleFadeInSeconds,
                    ["BubbleFadeOutSeconds"] = Overlay.BubbleFadeOutSeconds,
                    ["CommentDisplayIntervalSeconds"] = Overlay.CommentDisplayIntervalSeconds,
                    ["CommentDisplayIntervalMinSeconds"] = Overlay.CommentDisplayIntervalMinSeconds,
                    ["CommentDisplayIntervalMaxSeconds"] = Overlay.CommentDisplayIntervalMaxSeconds,
                    ["UseWhitespaceAwarePlacement"] = Overlay.UseWhitespaceAwarePlacement,
                    ["WhitespacePlacementProbability"] = Overlay.WhitespacePlacementProbability,
                    ["WhitespaceGridColumns"] = Overlay.WhitespaceGridColumns,
                    ["WhitespaceGridRows"] = Overlay.WhitespaceGridRows,
                    ["MaxSimultaneousComments"] = Overlay.MaxSimultaneousComments,
                    ["PresentationMonitorDeviceName"] = Overlay.PresentationMonitorDeviceName,
                    ["PresentationMonitorIndex"] = Overlay.PresentationMonitorIndex
                },
                ["Experiment"] = new Dictionary<string, object>
                {
                    ["EnableLogging"] = Experiment.EnableLogging,
                    ["SaveSlideTextToLog"] = Experiment.SaveSlideTextToLog,
                    ["SaveImagePathToLog"] = Experiment.SaveImagePathToLog
                },
                ["Preload"] = new Dictionary<string, object>
                {
                    ["EnablePreloadNextSlide"] = Preload.EnablePreloadNextSlide,
                    ["PreloadSlideCount"] = Preload.PreloadSlideCount,
                    ["PreloadCacheMaxSlides"] = Preload.PreloadCacheMaxSlides
                }
            };

            var serializer = new JavaScriptSerializer();
            File.WriteAllText(path, serializer.Serialize(root), Encoding.UTF8);
        }

        private void ApplyEnvironmentOverrides()
        {
            var model = Environment.GetEnvironmentVariable("GEMINI_MODEL");
            if (!string.IsNullOrWhiteSpace(model))
            {
                Gemini.Model = model.Trim();
            }
        }

        private void ReadGemini(Dictionary<string, object> root)
        {
            var section = ReadSection(root, "Gemini");
            Gemini.Model = ReadString(section, "Model", Gemini.Model);
            Gemini.ApiKeyEnvironmentVariable = ReadString(section, "ApiKeyEnvironmentVariable", Gemini.ApiKeyEnvironmentVariable);
            Gemini.TimeoutSeconds = ReadInt(section, "TimeoutSeconds", Gemini.TimeoutSeconds);
            Gemini.EnableApi = ReadBool(section, "EnableApi", Gemini.EnableApi);
        }

        private void ReadComments(Dictionary<string, object> root)
        {
            var section = ReadSection(root, "Comments");
            Comments.Mode = ReadString(section, "Mode", Comments.Mode);
            Comments.MaxCommentsPerSlide = ReadInt(section, "MaxCommentsPerSlide", Comments.MaxCommentsPerSlide);
            Comments.MinCharacters = ReadInt(section, "MinCharacters", Comments.MinCharacters);
            Comments.MaxCharacters = ReadInt(section, "MaxCharacters", Comments.MaxCharacters);
            Comments.Language = ReadString(section, "Language", Comments.Language);
        }

        private void ReadOverlay(Dictionary<string, object> root)
        {
            var section = ReadSection(root, "Overlay");
            Overlay.Position = ReadString(section, "Position", Overlay.Position);
            Overlay.Width = ReadDouble(section, "Width", Overlay.Width);
            Overlay.MarginRight = ReadDouble(section, "MarginRight", Overlay.MarginRight);
            Overlay.MarginBottom = ReadDouble(section, "MarginBottom", Overlay.MarginBottom);
            Overlay.FontSize = ReadDouble(section, "FontSize", Overlay.FontSize);
            Overlay.CommentFontSize = ReadDouble(section, "CommentFontSize", Overlay.CommentFontSize);
            Overlay.CommentTextColor = ReadString(section, "CommentTextColor", Overlay.CommentTextColor);
            Overlay.CommentBackgroundColor = ReadString(section, "CommentBackgroundColor", Overlay.CommentBackgroundColor);
            Overlay.UseNiconicoFlow = ReadBool(section, "UseNiconicoFlow", Overlay.UseNiconicoFlow);
            Overlay.DisplayMode = ReadEnum(section, "DisplayMode", Overlay.DisplayMode);
            Overlay.CommentLifetimeSeconds = ReadDouble(section, "CommentLifetimeSeconds", Overlay.CommentLifetimeSeconds);
            Overlay.FlowSpeedPixelsPerSecond = ReadDouble(section, "FlowSpeedPixelsPerSecond", Overlay.FlowSpeedPixelsPerSecond);
            Overlay.BubbleLifetimeSeconds = ReadDouble(section, "BubbleLifetimeSeconds", Overlay.BubbleLifetimeSeconds);
            Overlay.BubbleFadeSeconds = ReadDouble(section, "BubbleFadeSeconds", Overlay.BubbleFadeSeconds);
            Overlay.BubbleFadeInSeconds = ReadDouble(section, "BubbleFadeInSeconds", Overlay.BubbleFadeInSeconds);
            Overlay.BubbleFadeOutSeconds = ReadDouble(section, "BubbleFadeOutSeconds", Overlay.BubbleFadeOutSeconds);
            Overlay.CommentDisplayIntervalSeconds = ReadDouble(section, "CommentDisplayIntervalSeconds", Overlay.CommentDisplayIntervalSeconds);
            Overlay.CommentDisplayIntervalMinSeconds = ReadDouble(section, "CommentDisplayIntervalMinSeconds", Overlay.CommentDisplayIntervalMinSeconds);
            Overlay.CommentDisplayIntervalMaxSeconds = ReadDouble(section, "CommentDisplayIntervalMaxSeconds", Overlay.CommentDisplayIntervalMaxSeconds);
            Overlay.UseWhitespaceAwarePlacement = ReadBool(section, "UseWhitespaceAwarePlacement", Overlay.UseWhitespaceAwarePlacement);
            Overlay.WhitespacePlacementProbability = ReadDouble(section, "WhitespacePlacementProbability", Overlay.WhitespacePlacementProbability);
            Overlay.WhitespaceGridColumns = ReadInt(section, "WhitespaceGridColumns", Overlay.WhitespaceGridColumns);
            Overlay.WhitespaceGridRows = ReadInt(section, "WhitespaceGridRows", Overlay.WhitespaceGridRows);
            Overlay.MaxSimultaneousComments = ReadInt(section, "MaxSimultaneousComments", Overlay.MaxSimultaneousComments);
            Overlay.PresentationMonitorDeviceName = ReadString(section, "PresentationMonitorDeviceName", Overlay.PresentationMonitorDeviceName);
            Overlay.PresentationMonitorIndex = ReadInt(section, "PresentationMonitorIndex", Overlay.PresentationMonitorIndex);
        }

        private void ReadExperiment(Dictionary<string, object> root)
        {
            var section = ReadSection(root, "Experiment");
            Experiment.EnableLogging = ReadBool(section, "EnableLogging", Experiment.EnableLogging);
            Experiment.SaveSlideTextToLog = ReadBool(section, "SaveSlideTextToLog", Experiment.SaveSlideTextToLog);
            Experiment.SaveImagePathToLog = ReadBool(section, "SaveImagePathToLog", Experiment.SaveImagePathToLog);
        }

        private void ReadPreload(Dictionary<string, object> root)
        {
            var section = ReadSection(root, "Preload");
            Preload.EnablePreloadNextSlide = ReadBool(section, "EnablePreloadNextSlide", Preload.EnablePreloadNextSlide);
            Preload.PreloadSlideCount = ReadInt(section, "PreloadSlideCount", Preload.PreloadSlideCount);
            Preload.PreloadCacheMaxSlides = ReadInt(section, "PreloadCacheMaxSlides", Preload.PreloadCacheMaxSlides);
        }

        private static Dictionary<string, object> ReadSection(Dictionary<string, object> root, string key)
        {
            if (root != null && root.TryGetValue(key, out var value))
            {
                return value as Dictionary<string, object> ?? new Dictionary<string, object>();
            }

            return new Dictionary<string, object>();
        }

        private static string ReadString(Dictionary<string, object> section, string key, string fallback)
        {
            return section.TryGetValue(key, out var value) && value != null ? Convert.ToString(value) : fallback;
        }

        private static int ReadInt(Dictionary<string, object> section, string key, int fallback)
        {
            return section.TryGetValue(key, out var value) ? Convert.ToInt32(value) : fallback;
        }

        private static double ReadDouble(Dictionary<string, object> section, string key, double fallback)
        {
            return section.TryGetValue(key, out var value) ? Convert.ToDouble(value) : fallback;
        }

        private static bool ReadBool(Dictionary<string, object> section, string key, bool fallback)
        {
            return section.TryGetValue(key, out var value) ? Convert.ToBoolean(value) : fallback;
        }

        private static TEnum ReadEnum<TEnum>(Dictionary<string, object> section, string key, TEnum fallback)
            where TEnum : struct
        {
            if (!section.TryGetValue(key, out var value) || value == null)
            {
                return fallback;
            }

            return Enum.TryParse(Convert.ToString(value), true, out TEnum parsed)
                ? parsed
                : fallback;
        }
    }

    public class GeminiSettings
    {
        public string Model { get; set; } = "gemini-2.5-flash";

        public string ApiKeyEnvironmentVariable { get; set; } = "GEMINI_API_KEY";

        public int TimeoutSeconds { get; set; } = 20;

        public bool EnableApi { get; set; }
    }

    public class CommentSettings
    {
        public string Mode { get; set; } = "Mixed";

        public int MaxCommentsPerSlide { get; set; } = 10;

        public int MinCharacters { get; set; } = 10;

        public int MaxCharacters { get; set; } = 18;

        public string Language { get; set; } = "ja-JP";
    }

    public class OverlaySettings
    {
        public string Position { get; set; } = "BottomRight";

        public double Width { get; set; } = 420;

        public double MarginRight { get; set; } = 48;

        public double MarginBottom { get; set; } = 48;

        public double FontSize { get; set; } = 24;

        public double CommentFontSize { get; set; } = 28;

        public string CommentTextColor { get; set; } = "#FFFFFFFF";

        public string CommentBackgroundColor { get; set; } = "#99000000";

        public bool UseNiconicoFlow { get; set; }

        public OverlayDisplayMode DisplayMode { get; set; } = OverlayDisplayMode.Panel;

        public double CommentLifetimeSeconds { get; set; } = 7;

        public double FlowSpeedPixelsPerSecond { get; set; } = 80;

        public double BubbleLifetimeSeconds { get; set; } = 9;

        public double BubbleFadeSeconds { get; set; } = 1.0;

        public double CommentDisplayIntervalSeconds { get; set; } = 1.2;

        public double CommentDisplayIntervalMinSeconds { get; set; } = 1.0;

        public double CommentDisplayIntervalMaxSeconds { get; set; } = 2.5;

        public double BubbleFadeInSeconds { get; set; } = 1.0;

        public double BubbleFadeOutSeconds { get; set; } = 1.0;

        public bool UseWhitespaceAwarePlacement { get; set; } = true;

        public double WhitespacePlacementProbability { get; set; } = 0.7;

        public int WhitespaceGridColumns { get; set; } = 12;

        public int WhitespaceGridRows { get; set; } = 8;

        public int MaxSimultaneousComments { get; set; } = 3;

        public string PresentationMonitorDeviceName { get; set; } = string.Empty;

        public int PresentationMonitorIndex { get; set; } = -1;
    }

    public class ExperimentSettings
    {
        public bool EnableLogging { get; set; } = true;

        public bool SaveSlideTextToLog { get; set; }

        public bool SaveImagePathToLog { get; set; } = true;
    }

    public class PreloadSettings
    {
        public bool EnablePreloadNextSlide { get; set; } = true;

        public int PreloadSlideCount { get; set; } = 1;

        public int PreloadCacheMaxSlides { get; set; } = 5;
    }
}
