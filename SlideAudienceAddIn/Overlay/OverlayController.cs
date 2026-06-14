using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using SlideAudienceAddIn.Models;
using SlideAudienceAddIn.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideAudienceAddIn.Overlay
{
    public struct OverlayBounds
    {
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public string Source { get; set; }
        public long Hwnd { get; set; }
        public double DpiX { get; set; }
        public double DpiY { get; set; }

        public double Right => Left + Width;

        public double Bottom => Top + Height;
    }

    public class OverlayController
    {
        private readonly object _dispatcherLock = new object();
        private Dispatcher _dispatcher;
        private Thread _overlayThread;
        private OverlayWindow _window;
        private OverlaySettings _settings = new OverlaySettings();
        private int _slideSessionId;

        public void ApplySettings(OverlaySettings settings)
        {
            _settings = settings ?? new OverlaySettings();
            Debug.WriteLine($"[SlideAudience] OverlayController.ApplySettings display mode={_settings.DisplayMode}");
            Debug.WriteLine($"[SlideAudience] OverlayController.ApplySettings presentation monitor device={_settings.PresentationMonitorDeviceName}, index={_settings.PresentationMonitorIndex}");
            var dispatcher = _dispatcher;
            if (dispatcher != null && _window != null)
            {
                dispatcher.BeginInvoke(new Action(() => _window.ApplySettings(_settings)));
            }
        }

        public OverlayBounds? CaptureSlideShowBounds(PowerPoint.SlideShowWindow slideShowWindow)
        {
            LogAllScreens("CaptureSlideShowBounds");
            var hwnd = GetSlideShowHwnd(slideShowWindow);
            Debug.WriteLine($"[SlideAudience] SlideShowWindow.HWND=0x{hwnd.ToInt64():X}");
            if (hwnd != IntPtr.Zero && GetWindowRect(hwnd, out var rect))
            {
                var width = rect.Right - rect.Left;
                var height = rect.Bottom - rect.Top;
                Debug.WriteLine($"[SlideAudience] GetWindowRect succeeded hwnd=0x{hwnd.ToInt64():X}, Left={rect.Left}, Top={rect.Top}, Right={rect.Right}, Bottom={rect.Bottom}, Width={width}, Height={height}");
                if (width > 0 && height > 0)
                {
                    var dpi = GetDpiForWindowOrDefault(hwnd);
                    LogScreenForRect(rect.Left, rect.Top, width, height, "SlideShowWindow");
                    Debug.WriteLine($"[SlideAudience] SlideShowWindow DPI hwnd=0x{hwnd.ToInt64():X}, dpiX={dpi.Width:F1}, dpiY={dpi.Height:F1}");
                    return new OverlayBounds
                    {
                        Left = rect.Left,
                        Top = rect.Top,
                        Width = width,
                        Height = height,
                        Source = "SlideShowWindow.GetWindowRect",
                        Hwnd = hwnd.ToInt64(),
                        DpiX = dpi.Width,
                        DpiY = dpi.Height
                    };
                }

                Debug.WriteLine("[SlideAudience] GetWindowRect returned non-positive size; falling back to configured presentation monitor");
            }
            else
            {
                var error = hwnd == IntPtr.Zero ? 0 : Marshal.GetLastWin32Error();
                Debug.WriteLine($"[SlideAudience] GetWindowRect failed hwnd=0x{hwnd.ToInt64():X}, lastWin32Error={error}; falling back to configured presentation monitor");
            }

            return CaptureConfiguredPresentationMonitorBounds("SlideShowWindow HWND/rect unavailable");
        }

        private OverlayBounds? CaptureConfiguredPresentationMonitorBounds(string reason)
        {
            var screens = System.Windows.Forms.Screen.AllScreens;
            if (screens == null || screens.Length == 0)
            {
                Debug.WriteLine($"[SlideAudience] Presentation monitor fallback failed reason={reason}: Screen.AllScreens is empty");
                return null;
            }

            var screen = FindConfiguredPresentationMonitor(screens);
            if (screen == null)
            {
                var cursorScreen = System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position);
                screen = cursorScreen ?? screens[0];
                Debug.WriteLine($"[SlideAudience] Presentation monitor fallback using cursor screen reason={reason}, DeviceName={screen.DeviceName}");
            }
            else
            {
                Debug.WriteLine($"[SlideAudience] Presentation monitor fallback using configured screen reason={reason}, DeviceName={screen.DeviceName}, index={Array.IndexOf(screens, screen)}");
            }

            var bounds = screen.Bounds;
            Debug.WriteLine($"[SlideAudience] Presentation monitor fallback bounds DeviceName={screen.DeviceName}, Primary={screen.Primary}, Left={bounds.Left}, Top={bounds.Top}, Right={bounds.Right}, Bottom={bounds.Bottom}, Width={bounds.Width}, Height={bounds.Height}");
            return new OverlayBounds
            {
                Left = bounds.Left,
                Top = bounds.Top,
                Width = bounds.Width,
                Height = bounds.Height,
                Source = "ConfiguredPresentationMonitor",
                Hwnd = 0,
                DpiX = 96,
                DpiY = 96
            };
        }

        private System.Windows.Forms.Screen FindConfiguredPresentationMonitor(System.Windows.Forms.Screen[] screens)
        {
            if (screens == null || screens.Length == 0)
            {
                return null;
            }

            var configuredDeviceName = _settings.PresentationMonitorDeviceName;
            if (!string.IsNullOrWhiteSpace(configuredDeviceName))
            {
                var screen = screens.FirstOrDefault(candidate =>
                    string.Equals(candidate.DeviceName, configuredDeviceName, StringComparison.OrdinalIgnoreCase));
                if (screen != null)
                {
                    return screen;
                }

                Debug.WriteLine($"[SlideAudience] configured presentation monitor device not found: {configuredDeviceName}");
            }

            var configuredIndex = _settings.PresentationMonitorIndex;
            if (configuredIndex >= 0 && configuredIndex < screens.Length)
            {
                return screens[configuredIndex];
            }

            if (configuredIndex >= screens.Length)
            {
                Debug.WriteLine($"[SlideAudience] configured presentation monitor index out of range: {configuredIndex}, screenCount={screens.Length}");
            }

            return null;
        }

        private static void LogAllScreens(string context)
        {
            try
            {
                var screens = System.Windows.Forms.Screen.AllScreens;
                Debug.WriteLine($"[SlideAudience] Screen.AllScreens context={context}, count={screens.Length}");
                for (var i = 0; i < screens.Length; i++)
                {
                    var screen = screens[i];
                    var bounds = screen.Bounds;
                    var workArea = screen.WorkingArea;
                    Debug.WriteLine(
                        $"[SlideAudience] Screen.AllScreens[{i}] DeviceName={screen.DeviceName}, Primary={screen.Primary}, Bounds.Left={bounds.Left}, Bounds.Top={bounds.Top}, Bounds.Right={bounds.Right}, Bounds.Bottom={bounds.Bottom}, Bounds.Width={bounds.Width}, Bounds.Height={bounds.Height}, WorkingArea.Left={workArea.Left}, WorkingArea.Top={workArea.Top}, WorkingArea.Right={workArea.Right}, WorkingArea.Bottom={workArea.Bottom}, WorkingArea.Width={workArea.Width}, WorkingArea.Height={workArea.Height}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SlideAudience] failed to log Screen.AllScreens");
                Debug.WriteLine(ex.ToString());
            }
        }

        private static void LogScreenForRect(int left, int top, int width, int height, string label)
        {
            try
            {
                var rect = new System.Drawing.Rectangle(left, top, width, height);
                var screen = SelectScreenWithLargestIntersection(rect) ?? System.Windows.Forms.Screen.FromRectangle(rect);
                Debug.WriteLine($"[SlideAudience] {label} monitor DeviceName={screen.DeviceName}, Primary={screen.Primary}, Bounds.Left={screen.Bounds.Left}, Bounds.Top={screen.Bounds.Top}, Bounds.Right={screen.Bounds.Right}, Bounds.Bottom={screen.Bounds.Bottom}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] failed to resolve monitor for {label}");
                Debug.WriteLine(ex.ToString());
            }
        }

        private static System.Windows.Forms.Screen SelectScreenWithLargestIntersection(System.Drawing.Rectangle rect)
        {
            System.Windows.Forms.Screen bestScreen = null;
            long bestArea = 0;

            foreach (var screen in System.Windows.Forms.Screen.AllScreens)
            {
                var intersection = System.Drawing.Rectangle.Intersect(rect, screen.Bounds);
                var area = Math.Max(0, intersection.Width) * (long)Math.Max(0, intersection.Height);
                if (area > bestArea)
                {
                    bestArea = area;
                    bestScreen = screen;
                }
            }

            return bestScreen;
        }

        private static string FormatBounds(OverlayBounds bounds)
        {
            return $"Source={bounds.Source}, Hwnd=0x{bounds.Hwnd:X}, Left={bounds.Left:F0}, Top={bounds.Top:F0}, Right={bounds.Right:F0}, Bottom={bounds.Bottom:F0}, Width={bounds.Width:F0}, Height={bounds.Height:F0}, DpiX={bounds.DpiX:F1}, DpiY={bounds.DpiY:F1}";
        }

        private static Size ToDeviceIndependentSize(Window window, int width, int height)
        {
            var source = PresentationSource.FromVisual(window);
            if (source?.CompositionTarget == null)
            {
                return new Size(width, height);
            }

            var size = source.CompositionTarget.TransformFromDevice.Transform(new Vector(width, height));
            return new Size(Math.Max(1, size.X), Math.Max(1, size.Y));
        }

        public void Show(PowerPoint.SlideShowWindow slideShowWindow, int slideIndex)
        {
            Show(CaptureSlideShowBounds(slideShowWindow), slideIndex);
        }

        public void Show(OverlayBounds? bounds, int slideIndex)
        {
            ShowComments(bounds, CreateDummyComments(slideIndex), null);
        }

        public void ShowComments(PowerPoint.SlideShowWindow slideShowWindow, IEnumerable<string> comments)
        {
            ShowComments(CaptureSlideShowBounds(slideShowWindow), comments);
        }

        public void ShowComments(OverlayBounds? bounds, IEnumerable<string> comments)
        {
            var commentModels = (comments ?? new string[0])
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .Select(text => AudienceComment.Create("comment", text, persona: "audience"))
                .ToList();
            ShowComments(bounds, commentModels, null);
        }

        public void ShowComments(
            OverlayBounds? bounds,
            IReadOnlyList<AudienceComment> comments,
            IReadOnlyList<WhitespaceRegion> whitespaceRegions)
        {
            var sessionId = Volatile.Read(ref _slideSessionId);
            ShowComments(bounds, comments, whitespaceRegions, sessionId);
        }

        public void ShowComments(
            OverlayBounds? bounds,
            IReadOnlyList<AudienceComment> comments,
            IReadOnlyList<WhitespaceRegion> whitespaceRegions,
            int slideSessionId)
        {
            Debug.WriteLine($"[SlideAudience] OverlayController.ShowComments requested thread id={Thread.CurrentThread.ManagedThreadId}");
            Debug.WriteLine($"[SlideAudience] OverlayController.ShowComments sessionId={slideSessionId}, currentSessionId={Volatile.Read(ref _slideSessionId)}");
            if (!IsCurrentSession(slideSessionId))
            {
                Debug.WriteLine("[SlideAudience] OverlayController.ShowComments ignored stale slide session before dispatch");
                return;
            }

            var commentList = (comments ?? new List<AudienceComment>())
                .Where(comment => comment != null && !string.IsNullOrWhiteSpace(comment.Text))
                .ToList();
            var regionList = (whitespaceRegions ?? new List<WhitespaceRegion>())
                .Where(region => region != null)
                .ToList();
            Debug.WriteLine($"[SlideAudience] OverlayController.ShowComments comments queued count={commentList.Count}");
            Debug.WriteLine($"[SlideAudience] whitespace-aware placement enabled={_settings.UseWhitespaceAwarePlacement}, regions count={regionList.Count}");

            EnsureDispatcher();
            var dispatcher = _dispatcher;
            if (dispatcher == null)
            {
                Debug.WriteLine("[SlideAudience] OverlayController.ShowComments skipped because dispatcher is null");
                return;
            }

            try
            {
                dispatcher.BeginInvoke(new Action(() =>
                {
                    if (!IsCurrentSession(slideSessionId))
                    {
                        Debug.WriteLine("[SlideAudience] OverlayController.ShowComments ignored stale slide session on dispatcher");
                        return;
                    }

                    Debug.WriteLine($"[SlideAudience] OverlayController.ShowComments executing on dispatcher thread id={Thread.CurrentThread.ManagedThreadId}");
                    EnsureWindowOnDispatcher();
                    var targetBounds = ResolveBoundsOnDispatcher(bounds);
                    FitToBoundsOnDispatcher(targetBounds);

                    if (!_window.IsVisible)
                    {
                        Debug.WriteLine("[SlideAudience] OverlayController.ShowComments Window.Show called");
                        _window.Show();
                        FitToBoundsOnDispatcher(targetBounds);
                    }
                    else
                    {
                        Debug.WriteLine("[SlideAudience] OverlayController.ShowComments _window already visible");
                    }

                    _window.UpdateLayout();
                    Debug.WriteLine($"[SlideAudience] Overlay display update start bounds={FormatBounds(targetBounds)}, windowVisible={_window.IsVisible}");
                    Debug.WriteLine("[SlideAudience] OverlayController.ShowComments _window.SetComments called");
                    _window.SetComments(commentList, regionList);
                    Debug.WriteLine($"[SlideAudience] Overlay display update completed comments={commentList.Count}, regions={regionList.Count}, bounds={FormatBounds(targetBounds)}");

                }));
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SlideAudience] OverlayController.ShowComments dispatch failed");
                Debug.WriteLine(ex.ToString());
            }
        }

        public void UpdateComments(int slideIndex)
        {
            ShowComments((OverlayBounds?)null, CreateDummyComments(slideIndex), null);
        }

        public int StartNewSlideSession()
        {
            var sessionId = Interlocked.Increment(ref _slideSessionId);
            Debug.WriteLine($"[SlideAudience] OverlayController.StartNewSlideSession sessionId={sessionId}");
            return sessionId;
        }

        public bool IsCurrentSession(int sessionId)
        {
            return sessionId == Volatile.Read(ref _slideSessionId);
        }

        public void ResetForNewSlide(string reason = "slideChanged")
        {
            var dispatcher = _dispatcher;
            if (dispatcher == null)
            {
                Debug.WriteLine("[SlideAudience] OverlayController.ResetForNewSlide skipped because dispatcher is null");
                return;
            }

            var stopwatch = Stopwatch.StartNew();
            Action resetAction = () =>
            {
                Debug.WriteLine($"[SlideAudience] OverlayController.ResetForNewSlide executing reason={reason}");
                _window?.ResetForNewSlide(reason);
            };

            try
            {
                if (dispatcher.CheckAccess())
                {
                    resetAction();
                }
                else
                {
                    dispatcher.Invoke(resetAction);
                }

                stopwatch.Stop();
                Debug.WriteLine($"[SlideAudience] OverlayController.ResetForNewSlide completed in {stopwatch.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SlideAudience] OverlayController.ResetForNewSlide failed");
                Debug.WriteLine(ex.ToString());
            }
        }

        public void Clear(string reason = "manual")
        {
            var sessionId = Interlocked.Increment(ref _slideSessionId);
            Debug.WriteLine($"[SlideAudience] OverlayController.Clear invalidated slide session sessionId={sessionId}, reason={reason}");
            var dispatcher = _dispatcher;
            if (dispatcher == null)
            {
                return;
            }

            try
            {
                dispatcher.BeginInvoke(new Action(() =>
                {
                    _window?.ResetForNewSlide(reason);
                }));
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SlideAudience] OverlayController.Clear failed");
                Debug.WriteLine(ex.ToString());
            }
        }

        public void Close()
        {
            var sessionId = Interlocked.Increment(ref _slideSessionId);
            Debug.WriteLine($"[SlideAudience] OverlayController.Close invalidated slide session sessionId={sessionId}");
            var dispatcher = _dispatcher;
            if (dispatcher == null)
            {
                return;
            }

            Action closeAction = () =>
            {
                Debug.WriteLine($"[SlideAudience] OverlayController.Close executing on dispatcher thread id={Thread.CurrentThread.ManagedThreadId}");
                if (_window != null)
                {
                    var window = _window;
                    _window = null;
                    window.ResetForNewSlide("slideshowEnded");
                    Debug.WriteLine("[SlideAudience] OverlayController.Close window.Close called");
                    window.Close();
                }

                Dispatcher.CurrentDispatcher.BeginInvokeShutdown(DispatcherPriority.Background);
            };

            try
            {
                if (dispatcher.CheckAccess())
                {
                    closeAction();
                }
                else
                {
                    dispatcher.Invoke(closeAction);
                }

                lock (_dispatcherLock)
                {
                    _dispatcher = null;
                    _overlayThread = null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SlideAudience] OverlayController.Close failed");
                Debug.WriteLine(ex.ToString());
            }
        }

        private void EnsureDispatcher()
        {
            if (_dispatcher != null && !_dispatcher.HasShutdownStarted && !_dispatcher.HasShutdownFinished)
            {
                return;
            }

            lock (_dispatcherLock)
            {
                if (_dispatcher != null && !_dispatcher.HasShutdownStarted && !_dispatcher.HasShutdownFinished)
                {
                    return;
                }

                var ready = new ManualResetEventSlim(false);
                _overlayThread = new Thread(() =>
                {
                    Debug.WriteLine($"[SlideAudience] Overlay STA thread started thread id={Thread.CurrentThread.ManagedThreadId}");
                    _dispatcher = Dispatcher.CurrentDispatcher;
                    Debug.WriteLine($"[SlideAudience] Overlay Dispatcher created thread id={Thread.CurrentThread.ManagedThreadId}");
                    ready.Set();
                    Dispatcher.Run();
                    Debug.WriteLine($"[SlideAudience] Overlay Dispatcher stopped thread id={Thread.CurrentThread.ManagedThreadId}");
                });

                _overlayThread.Name = "SlideAudience Overlay STA";
                _overlayThread.IsBackground = true;
                _overlayThread.SetApartmentState(ApartmentState.STA);
                _overlayThread.Start();
                ready.Wait();
            }
        }

        private void EnsureWindowOnDispatcher()
        {
            if (_window != null)
            {
                return;
            }

            _window = new OverlayWindow();
            _window.WindowStartupLocation = WindowStartupLocation.Manual;
            _window.ApplySettings(_settings);
            Debug.WriteLine("[SlideAudience] OverlayWindow created");
        }

        private OverlayBounds ResolveBoundsOnDispatcher(OverlayBounds? bounds)
        {
            if (bounds.HasValue)
            {
                return bounds.Value;
            }

            return CaptureConfiguredPresentationMonitorBounds("overlay bounds missing")
                ?? new OverlayBounds
                {
                    Left = 0,
                    Top = 0,
                    Width = 1,
                    Height = 1,
                    Source = "EmergencyFallback",
                    Hwnd = 0,
                    DpiX = 96,
                    DpiY = 96
                };
        }

        private void FitToBoundsOnDispatcher(OverlayBounds bounds)
        {
            var left = (int)Math.Round(bounds.Left);
            var top = (int)Math.Round(bounds.Top);
            var width = Math.Max(1, (int)Math.Round(bounds.Width));
            var height = Math.Max(1, (int)Math.Round(bounds.Height));
            var helper = new WindowInteropHelper(_window);
            var hwnd = helper.Handle;
            if (hwnd == IntPtr.Zero)
            {
                hwnd = helper.EnsureHandle();
            }

            var dpiX = bounds.DpiX > 0 ? bounds.DpiX : GetDpiForWindowOrDefault(hwnd).Width;
            var dpiY = bounds.DpiY > 0 ? bounds.DpiY : GetDpiForWindowOrDefault(hwnd).Height;
            var dipSize = ToDeviceIndependentSize(_window, width, height);
            _window.Left = bounds.Left;
            _window.Top = bounds.Top;
            _window.Width = dipSize.Width;
            _window.Height = dipSize.Height;

            Debug.WriteLine($"[SlideAudience] OverlayWindow.HWND=0x{hwnd.ToInt64():X}");
            Debug.WriteLine($"[SlideAudience] Overlay viewport metrics dipWidth={dipSize.Width:F1}, dipHeight={dipSize.Height:F1}, physicalWidth={width}, physicalHeight={height}, dpiX={dpiX:F1}, dpiY={dpiY:F1}");
            Debug.WriteLine($"[SlideAudience] Overlay SetWindowPos requested hwnd=0x{hwnd.ToInt64():X}, Left={left}, Top={top}, Width={width}, Height={height}, bounds={FormatBounds(bounds)}");
            var moved = SetWindowPos(
                hwnd,
                HwndTopmost,
                left,
                top,
                width,
                height,
                SwpNoActivate | SwpNoOwnerZOrder);
            Debug.WriteLine($"[SlideAudience] Overlay SetWindowPos result={moved}, lastWin32Error={(moved ? 0 : Marshal.GetLastWin32Error())}");
        }

        private static IReadOnlyList<AudienceComment> CreateDummyComments(int slideIndex)
        {
            return new[]
            {
                AudienceComment.Create("understanding", $"Slide {slideIndex}要点は？", persona: "beginner"),
                AudienceComment.Create("interest", "ちょっと気になる", persona: "curious"),
                AudienceComment.Create("question", "なぜその表現？", persona: "skeptic")
            };
        }

        private static IntPtr GetSlideShowHwnd(PowerPoint.SlideShowWindow slideShowWindow)
        {
            try
            {
                return new IntPtr(slideShowWindow.HWND);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SlideAudience] OverlayController.GetSlideShowHwnd failed");
                Debug.WriteLine(ex.ToString());
                return IntPtr.Zero;
            }
        }

        private static readonly IntPtr HwndTopmost = new IntPtr(-1);
        private const uint SwpNoActivate = 0x0010;
        private const uint SwpNoOwnerZOrder = 0x0200;

        private static Size GetDpiForWindowOrDefault(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
            {
                return new Size(96, 96);
            }

            try
            {
                var dpi = GetDpiForWindow(hwnd);
                if (dpi > 0)
                {
                    return new Size(dpi, dpi);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SlideAudience] GetDpiForWindow failed; using 96 DPI fallback");
                Debug.WriteLine(ex.ToString());
            }

            return new Size(96, 96);
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(
            IntPtr hWnd,
            IntPtr hWndInsertAfter,
            int x,
            int y,
            int cx,
            int cy,
            uint flags);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    }
}
