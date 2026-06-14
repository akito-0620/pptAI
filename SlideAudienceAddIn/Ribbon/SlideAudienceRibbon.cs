using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using SlideAudienceAddIn.Models;

namespace SlideAudienceAddIn.Ribbon
{
    [ComVisible(true)]
    public class SlideAudienceRibbon : Office.IRibbonExtensibility
    {
        private static readonly string[] DisplayModes = { "Panel", "Flow", "Bubble" };
        private static readonly string[] Presets = { "None", "Panel", "Flow", "Bubble", "Debug" };
        private static readonly string[] MaxComments = { "3", "5", "10" };
        private static readonly string[] MaxSimultaneous = { "1", "2", "3" };
        private static readonly string[] IntervalValues = { "0.5", "1.0", "1.5", "2.0", "2.5", "3.0", "4.0", "5.0" };
        private static readonly string[] WhitespaceProbabilityValues = { "0.0", "0.25", "0.5", "0.7", "0.85", "1.0" };
        private static readonly string[] FontSizeValues = { "18", "22", "26", "28", "32", "36", "42", "48" };
        private static readonly string[] TextColorLabels = { "White", "Black", "Yellow", "Cyan", "Pink", "Green" };
        private static readonly string[] TextColorValues = { "#FFFFFFFF", "#FF000000", "#FFFFFF00", "#FF00FFFF", "#FFFF66CC", "#FF66FF66" };
        private static readonly string[] BackgroundColorLabels = { "Transparent", "Light Black", "Black", "Dark Black", "White", "Yellow" };
        private static readonly string[] BackgroundColorValues = { "#00000000", "#66000000", "#99000000", "#CC000000", "#99FFFFFF", "#99FFFF00" };
        private static readonly string[] FlowSpeedValues = { "40", "60", "80", "100", "120", "160", "200" };
        private static readonly string[] BubbleFadeValues = { "0.2", "0.5", "0.8", "1.0", "1.5", "2.0" };
        private static readonly string[] PreloadCounts = { "1", "2", "3" };

        private readonly ThisAddIn _addIn;
        private Office.IRibbonUI _ribbon;

        public SlideAudienceRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
        }

        public string GetCustomUI(string ribbonId)
        {
            return RibbonXml;
        }

        public void OnLoad(Office.IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            Debug.WriteLine("[SlideAudience] Ribbon loaded");
        }

        public void OnEnableToggle(Office.IRibbonControl control, bool pressed)
        {
            _addIn.IsEnabled = pressed;
            _ribbon?.Invalidate();
        }

        public bool GetEnablePressed(Office.IRibbonControl control)
        {
            return _addIn.IsEnabled;
        }

        public void OnApiToggle(Office.IRibbonControl control, bool pressed)
        {
            _addIn.SetEnableApi(pressed);
            if (pressed && !_addIn.HasGeminiApiKey())
            {
                MessageBox.Show(
                    "Gemini API key was not found. SlideAudience will use dummy fallback until the environment variable is set.",
                    "SlideAudience",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

            _ribbon?.Invalidate();
        }

        public bool GetApiPressed(Office.IRibbonControl control)
        {
            return _addIn.CurrentSettings?.Gemini?.EnableApi == true;
        }

        public void OnWhitespaceToggle(Office.IRibbonControl control, bool pressed)
        {
            _addIn.SetWhitespaceAwarePlacement(pressed);
            _ribbon?.Invalidate();
        }

        public bool GetWhitespacePressed(Office.IRibbonControl control)
        {
            return _addIn.CurrentSettings?.Overlay?.UseWhitespaceAwarePlacement == true;
        }

        public int GetPresentationMonitorCount(Office.IRibbonControl control)
        {
            return Screen.AllScreens.Length + 1;
        }

        public string GetPresentationMonitorLabel(Office.IRibbonControl control, int index)
        {
            if (index <= 0)
            {
                return "Auto";
            }

            var screens = Screen.AllScreens;
            var screenIndex = index - 1;
            if (screenIndex < 0 || screenIndex >= screens.Length)
            {
                return "Unavailable";
            }

            var screen = screens[screenIndex];
            var bounds = screen.Bounds;
            return $"{screenIndex}: {screen.DeviceName} {bounds.Width}x{bounds.Height} ({bounds.Left},{bounds.Top})";
        }

        public int GetPresentationMonitorSelectedIndex(Office.IRibbonControl control)
        {
            var overlay = _addIn.CurrentSettings?.Overlay;
            if (overlay == null)
            {
                return 0;
            }

            var screens = Screen.AllScreens;
            if (!string.IsNullOrWhiteSpace(overlay.PresentationMonitorDeviceName))
            {
                for (var i = 0; i < screens.Length; i++)
                {
                    if (string.Equals(screens[i].DeviceName, overlay.PresentationMonitorDeviceName, StringComparison.OrdinalIgnoreCase))
                    {
                        return i + 1;
                    }
                }
            }

            return overlay.PresentationMonitorIndex >= 0 && overlay.PresentationMonitorIndex < screens.Length
                ? overlay.PresentationMonitorIndex + 1
                : 0;
        }

        public void OnPresentationMonitorChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            var screens = Screen.AllScreens;
            var screenIndex = selectedIndex - 1;
            if (screenIndex < 0 || screenIndex >= screens.Length)
            {
                _addIn.SetPresentationMonitor(null, -1);
            }
            else
            {
                _addIn.SetPresentationMonitor(screens[screenIndex].DeviceName, screenIndex);
            }

            _ribbon?.Invalidate();
        }

        public int GetDisplayModeCount(Office.IRibbonControl control)
        {
            return DisplayModes.Length;
        }

        public string GetDisplayModeLabel(Office.IRibbonControl control, int index)
        {
            return DisplayModes[index];
        }

        public int GetDisplayModeSelectedIndex(Office.IRibbonControl control)
        {
            var current = _addIn.CurrentSettings?.Overlay?.DisplayMode.ToString() ?? "Panel";
            return IndexOf(DisplayModes, current, 0);
        }

        public void OnDisplayModeChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex < 0 || selectedIndex >= DisplayModes.Length)
            {
                return;
            }

            if (Enum.TryParse(DisplayModes[selectedIndex], out OverlayDisplayMode mode))
            {
                _addIn.SetDisplayMode(mode);
            }

            _ribbon?.Invalidate();
        }

        public int GetPresetCount(Office.IRibbonControl control)
        {
            return Presets.Length;
        }

        public string GetPresetLabel(Office.IRibbonControl control, int index)
        {
            return Presets[index];
        }

        public int GetPresetSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOf(Presets, _addIn.ConditionPreset, 0);
        }

        public void OnPresetChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex >= 0 && selectedIndex < Presets.Length)
            {
                _addIn.SetConditionPreset(Presets[selectedIndex]);
            }

            _ribbon?.Invalidate();
        }

        public int GetMaxCommentsCount(Office.IRibbonControl control)
        {
            return MaxComments.Length;
        }

        public string GetMaxCommentsLabel(Office.IRibbonControl control, int index)
        {
            return MaxComments[index];
        }

        public int GetMaxCommentsSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOf(MaxComments, (_addIn.CurrentSettings?.Comments?.MaxCommentsPerSlide ?? 10).ToString(), 2);
        }

        public void OnMaxCommentsChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex >= 0 && selectedIndex < MaxComments.Length && int.TryParse(MaxComments[selectedIndex], out var value))
            {
                _addIn.SetMaxCommentsPerSlide(value);
            }

            _ribbon?.Invalidate();
        }

        public int GetMaxSimultaneousCount(Office.IRibbonControl control)
        {
            return MaxSimultaneous.Length;
        }

        public string GetMaxSimultaneousLabel(Office.IRibbonControl control, int index)
        {
            return MaxSimultaneous[index];
        }

        public int GetMaxSimultaneousSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOf(MaxSimultaneous, (_addIn.CurrentSettings?.Overlay?.MaxSimultaneousComments ?? 3).ToString(), 2);
        }

        public void OnMaxSimultaneousChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex >= 0 && selectedIndex < MaxSimultaneous.Length && int.TryParse(MaxSimultaneous[selectedIndex], out var value))
            {
                _addIn.SetMaxSimultaneousComments(value);
            }

            _ribbon?.Invalidate();
        }

        public int GetIntervalCount(Office.IRibbonControl control)
        {
            return IntervalValues.Length;
        }

        public string GetIntervalLabel(Office.IRibbonControl control, int index)
        {
            return IntervalValues[index];
        }

        public int GetIntervalMinSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOfDouble(IntervalValues, _addIn.CurrentSettings?.Overlay?.CommentDisplayIntervalMinSeconds ?? 1.0, 1);
        }

        public int GetIntervalMaxSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOfDouble(IntervalValues, _addIn.CurrentSettings?.Overlay?.CommentDisplayIntervalMaxSeconds ?? 2.5, 4);
        }

        public void OnIntervalMinChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (TryGetDouble(IntervalValues, selectedIndex, out var value))
            {
                _addIn.SetCommentDisplayIntervalMinSeconds(value);
            }

            _ribbon?.Invalidate();
        }

        public void OnIntervalMaxChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (TryGetDouble(IntervalValues, selectedIndex, out var value))
            {
                _addIn.SetCommentDisplayIntervalMaxSeconds(value);
            }

            _ribbon?.Invalidate();
        }

        public int GetWhitespaceProbabilityCount(Office.IRibbonControl control)
        {
            return WhitespaceProbabilityValues.Length;
        }

        public string GetWhitespaceProbabilityLabel(Office.IRibbonControl control, int index)
        {
            return WhitespaceProbabilityValues[index];
        }

        public int GetWhitespaceProbabilitySelectedIndex(Office.IRibbonControl control)
        {
            return IndexOfDouble(WhitespaceProbabilityValues, _addIn.CurrentSettings?.Overlay?.WhitespacePlacementProbability ?? 0.7, 3);
        }

        public void OnWhitespaceProbabilityChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (TryGetDouble(WhitespaceProbabilityValues, selectedIndex, out var value))
            {
                _addIn.SetWhitespacePlacementProbability(value);
            }

            _ribbon?.Invalidate();
        }

        public int GetFontSizeCount(Office.IRibbonControl control)
        {
            return FontSizeValues.Length;
        }

        public string GetFontSizeLabel(Office.IRibbonControl control, int index)
        {
            return FontSizeValues[index];
        }

        public int GetFontSizeSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOfDouble(FontSizeValues, _addIn.CurrentSettings?.Overlay?.CommentFontSize ?? 28, 3);
        }

        public void OnFontSizeChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (TryGetDouble(FontSizeValues, selectedIndex, out var value))
            {
                _addIn.SetCommentFontSize(value);
            }

            _ribbon?.Invalidate();
        }

        public int GetTextColorCount(Office.IRibbonControl control)
        {
            return TextColorLabels.Length;
        }

        public string GetTextColorLabel(Office.IRibbonControl control, int index)
        {
            return TextColorLabels[index];
        }

        public int GetTextColorSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOf(TextColorValues, _addIn.CurrentSettings?.Overlay?.CommentTextColor ?? "#FFFFFFFF", 0);
        }

        public void OnTextColorChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex >= 0 && selectedIndex < TextColorValues.Length)
            {
                _addIn.SetCommentTextColor(TextColorValues[selectedIndex]);
            }

            _ribbon?.Invalidate();
        }

        public int GetBackgroundColorCount(Office.IRibbonControl control)
        {
            return BackgroundColorLabels.Length;
        }

        public string GetBackgroundColorLabel(Office.IRibbonControl control, int index)
        {
            return BackgroundColorLabels[index];
        }

        public int GetBackgroundColorSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOf(BackgroundColorValues, _addIn.CurrentSettings?.Overlay?.CommentBackgroundColor ?? "#99000000", 2);
        }

        public void OnBackgroundColorChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex >= 0 && selectedIndex < BackgroundColorValues.Length)
            {
                _addIn.SetCommentBackgroundColor(BackgroundColorValues[selectedIndex]);
            }

            _ribbon?.Invalidate();
        }

        public int GetFlowSpeedCount(Office.IRibbonControl control)
        {
            return FlowSpeedValues.Length;
        }

        public string GetFlowSpeedLabel(Office.IRibbonControl control, int index)
        {
            return FlowSpeedValues[index];
        }

        public int GetFlowSpeedSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOfDouble(FlowSpeedValues, _addIn.CurrentSettings?.Overlay?.FlowSpeedPixelsPerSecond ?? 80, 2);
        }

        public void OnFlowSpeedChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (TryGetDouble(FlowSpeedValues, selectedIndex, out var value))
            {
                _addIn.SetFlowSpeedPixelsPerSecond(value);
            }

            _ribbon?.Invalidate();
        }

        public int GetBubbleFadeCount(Office.IRibbonControl control)
        {
            return BubbleFadeValues.Length;
        }

        public string GetBubbleFadeLabel(Office.IRibbonControl control, int index)
        {
            return BubbleFadeValues[index];
        }

        public int GetBubbleFadeInSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOfDouble(BubbleFadeValues, _addIn.CurrentSettings?.Overlay?.BubbleFadeInSeconds ?? 1.0, 3);
        }

        public int GetBubbleFadeOutSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOfDouble(BubbleFadeValues, _addIn.CurrentSettings?.Overlay?.BubbleFadeOutSeconds ?? 1.0, 3);
        }

        public void OnBubbleFadeInChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (TryGetDouble(BubbleFadeValues, selectedIndex, out var value))
            {
                _addIn.SetBubbleFadeInSeconds(value);
            }

            _ribbon?.Invalidate();
        }

        public void OnBubbleFadeOutChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (TryGetDouble(BubbleFadeValues, selectedIndex, out var value))
            {
                _addIn.SetBubbleFadeOutSeconds(value);
            }

            _ribbon?.Invalidate();
        }

        public void OnPreloadToggle(Office.IRibbonControl control, bool pressed)
        {
            _addIn.SetPreloadNextSlide(pressed);
            _ribbon?.Invalidate();
        }

        public bool GetPreloadPressed(Office.IRibbonControl control)
        {
            return _addIn.CurrentSettings?.Preload?.EnablePreloadNextSlide == true;
        }

        public void ClearPreloadCache(Office.IRibbonControl control)
        {
            _addIn.ClearPreloadCache();
            _ribbon?.Invalidate();
        }

        public int GetPreloadCountCount(Office.IRibbonControl control)
        {
            return PreloadCounts.Length;
        }

        public string GetPreloadCountLabel(Office.IRibbonControl control, int index)
        {
            return PreloadCounts[index];
        }

        public int GetPreloadCountSelectedIndex(Office.IRibbonControl control)
        {
            return IndexOf(PreloadCounts, (_addIn.CurrentSettings?.Preload?.PreloadSlideCount ?? 1).ToString(), 0);
        }

        public void OnPreloadCountChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex >= 0 && selectedIndex < PreloadCounts.Length && int.TryParse(PreloadCounts[selectedIndex], out var value))
            {
                _addIn.SetPreloadSlideCount(value);
            }

            _ribbon?.Invalidate();
        }

        public void TestCurrentSlide(Office.IRibbonControl control)
        {
            _addIn.TestCurrentSlide();
        }

        public void ClearOverlay(Office.IRibbonControl control)
        {
            _addIn.ClearOverlay();
        }

        public void OpenLogFolder(Office.IRibbonControl control)
        {
            _addIn.OpenLogFolder();
        }

        private static int IndexOf(string[] values, string value, int fallback)
        {
            for (var i = 0; i < values.Length; i++)
            {
                if (string.Equals(values[i], value, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return fallback;
        }

        private static int IndexOfDouble(string[] values, double value, int fallback)
        {
            for (var i = 0; i < values.Length; i++)
            {
                if (double.TryParse(values[i], NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) &&
                    Math.Abs(parsed - value) < 0.001)
                {
                    return i;
                }
            }

            return fallback;
        }

        private static bool TryGetDouble(string[] values, int index, out double value)
        {
            value = 0;
            return index >= 0 &&
                index < values.Length &&
                double.TryParse(values[index], NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private const string RibbonXml =
@"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnLoad"">
  <ribbon>
    <tabs>
      <tab id=""SlideAudienceTab"" label=""SlideAudience"">
        <group id=""RunGroup"" label=""Run"">
          <toggleButton id=""EnableToggle"" label=""Enable"" imageMso=""HappyFace"" size=""large"" getPressed=""GetEnablePressed"" onAction=""OnEnableToggle"" />
          <button id=""TestCurrentSlideButton"" label=""Test Current Slide"" imageMso=""ReviewNewComment"" size=""large"" onAction=""TestCurrentSlide"" />
          <button id=""ClearOverlayButton"" label=""Clear Overlay"" imageMso=""Delete"" onAction=""ClearOverlay"" />
          <button id=""OpenLogFolderButton"" label=""Open Log Folder"" imageMso=""OpenFolder"" onAction=""OpenLogFolder"" />
        </group>
        <group id=""ConditionGroup"" label=""Condition"">
          <dropDown id=""PresetDropdown"" label=""Preset"" getItemCount=""GetPresetCount"" getItemLabel=""GetPresetLabel"" getSelectedItemIndex=""GetPresetSelectedIndex"" onAction=""OnPresetChanged"" />
          <dropDown id=""DisplayModeDropdown"" label=""Display Mode"" getItemCount=""GetDisplayModeCount"" getItemLabel=""GetDisplayModeLabel"" getSelectedItemIndex=""GetDisplayModeSelectedIndex"" onAction=""OnDisplayModeChanged"" />
        </group>
        <group id=""GenerationGroup"" label=""Generation"">
          <toggleButton id=""ApiToggle"" label=""Gemini API"" imageMso=""HyperlinkOpen"" getPressed=""GetApiPressed"" onAction=""OnApiToggle"" />
          <dropDown id=""MaxCommentsDropdown"" label=""Max Comments"" getItemCount=""GetMaxCommentsCount"" getItemLabel=""GetMaxCommentsLabel"" getSelectedItemIndex=""GetMaxCommentsSelectedIndex"" onAction=""OnMaxCommentsChanged"" />
        </group>
        <group id=""PlacementGroup"" label=""Placement"">
          <toggleButton id=""WhitespaceToggle"" label=""Whitespace"" imageMso=""ObjectAlignMenu"" getPressed=""GetWhitespacePressed"" onAction=""OnWhitespaceToggle"" />
          <dropDown id=""PresentationMonitorDropdown"" label=""Monitor"" getItemCount=""GetPresentationMonitorCount"" getItemLabel=""GetPresentationMonitorLabel"" getSelectedItemIndex=""GetPresentationMonitorSelectedIndex"" onAction=""OnPresentationMonitorChanged"" />
          <dropDown id=""MaxSimultaneousDropdown"" label=""Max Visible"" getItemCount=""GetMaxSimultaneousCount"" getItemLabel=""GetMaxSimultaneousLabel"" getSelectedItemIndex=""GetMaxSimultaneousSelectedIndex"" onAction=""OnMaxSimultaneousChanged"" />
        </group>
        <group id=""TimingGroup"" label=""Timing"">
          <dropDown id=""IntervalMinDropdown"" label=""Interval Min"" getItemCount=""GetIntervalCount"" getItemLabel=""GetIntervalLabel"" getSelectedItemIndex=""GetIntervalMinSelectedIndex"" onAction=""OnIntervalMinChanged"" />
          <dropDown id=""IntervalMaxDropdown"" label=""Interval Max"" getItemCount=""GetIntervalCount"" getItemLabel=""GetIntervalLabel"" getSelectedItemIndex=""GetIntervalMaxSelectedIndex"" onAction=""OnIntervalMaxChanged"" />
        </group>
        <group id=""AppearanceGroup"" label=""Appearance"">
          <dropDown id=""FontSizeDropdown"" label=""Font Size"" getItemCount=""GetFontSizeCount"" getItemLabel=""GetFontSizeLabel"" getSelectedItemIndex=""GetFontSizeSelectedIndex"" onAction=""OnFontSizeChanged"" />
          <dropDown id=""TextColorDropdown"" label=""Text Color"" getItemCount=""GetTextColorCount"" getItemLabel=""GetTextColorLabel"" getSelectedItemIndex=""GetTextColorSelectedIndex"" onAction=""OnTextColorChanged"" />
          <dropDown id=""BackgroundColorDropdown"" label=""Background"" getItemCount=""GetBackgroundColorCount"" getItemLabel=""GetBackgroundColorLabel"" getSelectedItemIndex=""GetBackgroundColorSelectedIndex"" onAction=""OnBackgroundColorChanged"" />
        </group>
        <group id=""MotionGroup"" label=""Motion"">
          <dropDown id=""WhitespaceProbabilityDropdown"" label=""Whitespace Prob"" getItemCount=""GetWhitespaceProbabilityCount"" getItemLabel=""GetWhitespaceProbabilityLabel"" getSelectedItemIndex=""GetWhitespaceProbabilitySelectedIndex"" onAction=""OnWhitespaceProbabilityChanged"" />
          <dropDown id=""FlowSpeedDropdown"" label=""Flow Speed"" getItemCount=""GetFlowSpeedCount"" getItemLabel=""GetFlowSpeedLabel"" getSelectedItemIndex=""GetFlowSpeedSelectedIndex"" onAction=""OnFlowSpeedChanged"" />
          <dropDown id=""BubbleFadeInDropdown"" label=""Bubble In"" getItemCount=""GetBubbleFadeCount"" getItemLabel=""GetBubbleFadeLabel"" getSelectedItemIndex=""GetBubbleFadeInSelectedIndex"" onAction=""OnBubbleFadeInChanged"" />
          <dropDown id=""BubbleFadeOutDropdown"" label=""Bubble Out"" getItemCount=""GetBubbleFadeCount"" getItemLabel=""GetBubbleFadeLabel"" getSelectedItemIndex=""GetBubbleFadeOutSelectedIndex"" onAction=""OnBubbleFadeOutChanged"" />
        </group>
        <group id=""PreloadGroup"" label=""Preload"">
          <toggleButton id=""PreloadToggle"" label=""Preload Next"" imageMso=""Refresh"" getPressed=""GetPreloadPressed"" onAction=""OnPreloadToggle"" />
          <button id=""ClearPreloadCacheButton"" label=""Clear Cache"" imageMso=""Clear"" onAction=""ClearPreloadCache"" />
          <dropDown id=""PreloadCountDropdown"" label=""Preload Count"" getItemCount=""GetPreloadCountCount"" getItemLabel=""GetPreloadCountLabel"" getSelectedItemIndex=""GetPreloadCountSelectedIndex"" onAction=""OnPreloadCountChanged"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }
}
