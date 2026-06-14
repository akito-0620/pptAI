using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using SlideAudienceAddIn.Models;
using SlideAudienceAddIn.Services;

namespace SlideAudienceAddIn.Overlay
{
    public partial class OverlayWindow : Window
    {
        private readonly Random _random = new Random();
        private readonly Queue<QueuedComment> _commentQueue = new Queue<QueuedComment>();
        private readonly List<ActiveComment> _activeComments = new List<ActiveComment>();
        private readonly List<DispatcherTimer> _lifecycleTimers = new List<DispatcherTimer>();
        private readonly List<double> _recentFlowTops = new List<double>();
        private readonly List<Point> _recentBubblePositions = new List<Point>();
        private readonly List<int> _recentRegionIndexes = new List<int>();
        private readonly string[] _fallbackAnchors = { "rightTop", "rightMiddle", "rightBottom", "centerTop", "centerBottom" };
        private OverlaySettings _settings = new OverlaySettings();
        private DispatcherTimer _panelHideTimer;
        private DispatcherTimer _displayTimer;
        private IReadOnlyList<WhitespaceRegion> _currentWhitespaceRegions = new List<WhitespaceRegion>();
        private OverlayDisplayMode _currentDisplayMode = OverlayDisplayMode.Panel;
        private DateTime _lastDisplayUtc = DateTime.MinValue;
        private string _lastFallbackAnchor;
        private int _lastFlowLane = -1;
        private int _displaySessionId;
        private int _shownCount;
        private DateTime _nextDisplayDueUtc = DateTime.MinValue;
        private double _lastScheduledDelaySeconds;

        public OverlayWindow()
        {
            InitializeComponent();
            SourceInitialized += OnSourceInitialized;
            ApplySettings(_settings);
        }

        public void ApplySettings(OverlaySettings settings)
        {
            _settings = settings ?? new OverlaySettings();
            Debug.WriteLine($"[SlideAudience] displayMode={_settings.DisplayMode}");
            Debug.WriteLine($"[SlideAudience] maxSimultaneousComments={GetMaxSimultaneousComments()}");
            Debug.WriteLine($"[SlideAudience] CommentDisplayIntervalMinSeconds={GetDisplayIntervalMinSeconds():F1}");
            Debug.WriteLine($"[SlideAudience] CommentDisplayIntervalMaxSeconds={GetDisplayIntervalMaxSeconds():F1}");
            Debug.WriteLine($"[SlideAudience] whitespace-aware placement enabled={_settings.UseWhitespaceAwarePlacement}");
            Debug.WriteLine($"[SlideAudience] WhitespacePlacementProbability={GetWhitespacePlacementProbability():F2}");
            Debug.WriteLine($"[SlideAudience] CommentFontSize={GetCommentFontSize():F1}");
            Debug.WriteLine($"[SlideAudience] FlowSpeedPixelsPerSecond={GetFlowSpeedPixelsPerSecond():F1}");
            Debug.WriteLine($"[SlideAudience] BubbleFadeInSeconds={GetBubbleFadeInSeconds():F1}");
            Debug.WriteLine($"[SlideAudience] BubbleFadeOutSeconds={GetBubbleFadeOutSeconds():F1}");

            PanelBorder.Width = _settings.Width;
            PanelBorder.Margin = new Thickness(0, 0, _settings.MarginRight, _settings.MarginBottom);
            PanelBorder.Background = ParseBrushOrFallback(_settings.CommentBackgroundColor, "#99000000");
        }

        public void SetComments(IEnumerable<string> comments)
        {
            var commentModels = (comments ?? new string[0])
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .Select(text => AudienceComment.Create("comment", text, persona: "audience"))
                .ToList();
            SetComments(commentModels, null);
        }

        public void SetComments(
            IReadOnlyList<AudienceComment> comments,
            IReadOnlyList<WhitespaceRegion> whitespaceRegions)
        {
            var commentList = (comments ?? new List<AudienceComment>())
                .Where(comment => comment != null && !string.IsNullOrWhiteSpace(comment.Text))
                .ToList();
            var regionList = (whitespaceRegions ?? new List<WhitespaceRegion>())
                .Where(region => region != null)
                .OrderByDescending(region => region.Score)
                .ToList();

            Action updateAction = () =>
            {
                ClearComments("slideChanged");
                _displaySessionId++;
                _shownCount = 0;
                _lastDisplayUtc = DateTime.MinValue;
                _nextDisplayDueUtc = DateTime.MinValue;
                _lastScheduledDelaySeconds = 0;
                _lastFlowLane = -1;
                _lastFallbackAnchor = null;
                _currentWhitespaceRegions = regionList;
                _currentDisplayMode = _settings.DisplayMode;
                _recentFlowTops.Clear();
                _recentBubblePositions.Clear();
                _recentRegionIndexes.Clear();

                for (var i = 0; i < commentList.Count; i++)
                {
                    _commentQueue.Enqueue(new QueuedComment
                    {
                        Comment = commentList[i],
                        CommentIndex = i,
                        DisplayOrder = i
                    });
                }

                Debug.WriteLine($"[SlideAudience] SetComments displayMode={_settings.DisplayMode}");
                Debug.WriteLine($"[SlideAudience] queued comments count={_commentQueue.Count}");
                Debug.WriteLine($"[SlideAudience] active comments count={_activeComments.Count}");
                Debug.WriteLine($"[SlideAudience] maxSimultaneousComments={GetMaxSimultaneousComments()}");
                Debug.WriteLine($"[SlideAudience] detected whitespace region count={regionList.Count}");

                if (_settings.DisplayMode == OverlayDisplayMode.Panel)
                {
                    ShowPanelComments(commentList);
                    return;
                }

                PanelBorder.Visibility = Visibility.Collapsed;
                AnimationCanvas.Children.Clear();
                TryDisplayNextComment(_displaySessionId, forceImmediate: true);
                StartDisplayTimer(_displaySessionId);
            };

            if (Dispatcher.CheckAccess())
            {
                updateAction();
            }
            else
            {
                Dispatcher.Invoke(updateAction);
            }
        }

        public void ClearComments(string reason)
        {
            ResetForNewSlide(reason);
        }

        public void ResetForNewSlide(string reason = "slideChanged")
        {
            Action resetAction = () => ClearAllComments(reason);
            if (Dispatcher.CheckAccess())
            {
                resetAction();
            }
            else
            {
                Dispatcher.BeginInvoke(resetAction);
            }
        }

        public void ClearAllComments(string reason)
        {
            var stopwatch = Stopwatch.StartNew();
            _displaySessionId++;
            var removedCount = _activeComments.Count;

            StopAllTimers();
            StopAllAnimations(reason);

            _commentQueue.Clear();
            _activeComments.Clear();
            _recentFlowTops.Clear();
            _recentBubblePositions.Clear();
            _recentRegionIndexes.Clear();
            _lastDisplayUtc = DateTime.MinValue;
            _nextDisplayDueUtc = DateTime.MinValue;
            _lastScheduledDelaySeconds = 0;
            _shownCount = 0;

            AnimationCanvas.Children.Clear();
            CommentPanel.Children.Clear();
            PanelBorder.Visibility = Visibility.Collapsed;

            stopwatch.Stop();
            Debug.WriteLine($"[SlideAudience] Overlay reset took {stopwatch.ElapsedMilliseconds} ms, removed {removedCount} comments, reason={reason}");
            Debug.WriteLine($"[SlideAudience] active comments count={_activeComments.Count}");
            Debug.WriteLine($"[SlideAudience] canvas children count={AnimationCanvas.Children.Count}");
        }

        private void ShowPanelComments(IReadOnlyList<AudienceComment> comments)
        {
            StopPanelHideTimer();
            AnimationCanvas.Children.Clear();
            PanelBorder.Visibility = Visibility.Visible;
            PanelBorder.Opacity = 1;
            CommentPanel.Children.Clear();

            var shown = comments.Take(3).ToList();
            Debug.WriteLine($"[SlideAudience] comments shown count={shown.Count}");
            Debug.WriteLine($"[SlideAudience] active comments count={shown.Count}");
            Debug.WriteLine($"[SlideAudience] canvas children count={AnimationCanvas.Children.Count}");

            foreach (var comment in shown)
            {
                CommentPanel.Children.Add(CreateCommentRow(comment.Text));
            }

            ResetPanelHideTimer();
        }

        private void StartDisplayTimer(int sessionId)
        {
            StopDisplayTimer();
            _displayTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(200)
            };
            _displayTimer.Tick += (sender, args) =>
            {
                Debug.WriteLine($"[SlideAudience] timer tick sessionId={sessionId}, active comments count={_activeComments.Count}, queued comments count={_commentQueue.Count}");
                TryDisplayNextComment(sessionId, forceImmediate: false);
            };
            _displayTimer.Start();
        }

        private void TryDisplayNextComment(int sessionId, bool forceImmediate)
        {
            if (sessionId != _displaySessionId)
            {
                Debug.WriteLine($"[SlideAudience] displaySessionId mismatch ignored old={sessionId}, current={_displaySessionId}");
                return;
            }

            if (_commentQueue.Count == 0)
            {
                Debug.WriteLine("[SlideAudience] queue empty");
                StopDisplayTimer();
                return;
            }

            if (_activeComments.Count >= GetMaxSimultaneousComments())
            {
                Debug.WriteLine("[SlideAudience] wait because activeCommentsCount >= MaxSimultaneousComments");
                return;
            }

            if (!forceImmediate && DateTime.UtcNow < _nextDisplayDueUtc)
            {
                return;
            }

            var queued = _commentQueue.Dequeue();
            _lastDisplayUtc = DateTime.UtcNow;
            _lastScheduledDelaySeconds = GetNextCommentDelaySeconds();
            _nextDisplayDueUtc = _lastDisplayUtc.AddSeconds(_lastScheduledDelaySeconds);
            Debug.WriteLine($"[SlideAudience] dequeue comment index={queued.CommentIndex}, displayOrder={queued.DisplayOrder}, remaining={_commentQueue.Count}");
            Debug.WriteLine($"[SlideAudience] next comment scheduledDelaySeconds={_lastScheduledDelaySeconds:F2}");

            if (_currentDisplayMode == OverlayDisplayMode.Flow)
            {
                CreateFlowItem(queued, sessionId);
            }
            else if (_currentDisplayMode == OverlayDisplayMode.Bubble)
            {
                CreateBubbleItem(queued, sessionId);
            }
        }

        private UIElement CreateCommentRow(string comment)
        {
            var marker = new Border
            {
                Width = 5,
                Height = 30,
                CornerRadius = new CornerRadius(3),
                Background = (Brush)FindResource("AccentBrush"),
                Margin = new Thickness(0, 3, 14, 3)
            };

            var text = new TextBlock
            {
                Text = comment,
                Foreground = ParseBrushOrFallback(_settings.CommentTextColor, "#FFFFFFFF"),
                FontSize = GetCommentFontSize(),
                FontWeight = FontWeights.SemiBold,
                TextWrapping = TextWrapping.Wrap,
                LineHeight = GetCommentFontSize() + 8
            };

            var row = new DockPanel
            {
                LastChildFill = true,
                Margin = new Thickness(0, 0, 0, 12)
            };
            DockPanel.SetDock(marker, Dock.Left);
            row.Children.Add(marker);
            row.Children.Add(text);
            return row;
        }

        private Border CreateAnimatedComment(AudienceComment comment)
        {
            return new Border
            {
                Padding = new Thickness(16, 10, 16, 10),
                Background = ParseBrushOrFallback(_settings.CommentBackgroundColor, "#99000000"),
                CornerRadius = new CornerRadius(8),
                Child = new TextBlock
                {
                    Text = comment.Text,
                    Foreground = ParseBrushOrFallback(_settings.CommentTextColor, "#FFFFFFFF"),
                    FontSize = GetCommentFontSize(),
                    FontWeight = FontWeights.SemiBold,
                    TextWrapping = TextWrapping.NoWrap
                }
            };
        }

        private void CreateFlowItem(QueuedComment queued, int sessionId)
        {
            var item = CreateAnimatedComment(queued.Comment);
            AnimationCanvas.Children.Add(item);
            item.UpdateLayout();

            var laneIndex = SelectFlowLane();
            var estimatedWidth = EstimateCommentWidth(queued.Comment.Text, item);
            var margin = 80.0;
            var startX = ActualWidth + margin;
            var endX = -estimatedWidth - margin;
            var speed = GetFlowSpeedPixelsPerSecond();
            var distance = startX - endX;
            var durationSeconds = Math.Max(8, distance / speed);
            var placement = SelectFlowPlacement(laneIndex, _currentWhitespaceRegions);

            Canvas.SetLeft(item, startX);
            Canvas.SetTop(item, placement.Top);

            var active = new ActiveComment
            {
                Comment = queued.Comment,
                CommentIndex = queued.CommentIndex,
                DisplayOrder = queued.DisplayOrder,
                LaneOrSlotIndex = laneIndex,
                Element = item,
                Mode = OverlayDisplayMode.Flow,
                Bounds = new Rect(startX, placement.Top, estimatedWidth, Math.Max(48, item.ActualHeight))
            };
            _activeComments.Add(active);
            _shownCount++;

            Debug.WriteLine($"[SlideAudience] Flow lane selected={laneIndex}");
            Debug.WriteLine($"[SlideAudience] Flow lane index={laneIndex}");
            Debug.WriteLine($"[SlideAudience] Flow startX={startX:F1}, endX={endX:F1}, estimatedWidth={estimatedWidth:F1}, durationSeconds={durationSeconds:F2}, ActualWidth={ActualWidth:F1}");
            Debug.WriteLine($"[SlideAudience] Flow item created: {queued.Comment.Text}, commentIndex={queued.CommentIndex}, activeCommentsCount={_activeComments.Count}, queueRemaining={_commentQueue.Count}");
            Debug.WriteLine($"[SlideAudience] canvas children count={AnimationCanvas.Children.Count}");
            Debug.WriteLine($"[SlideAudience] selected region for each comment index={queued.CommentIndex}, mode={placement.Mode}, regionIndex={placement.RegionIndex}, top={placement.Top:F1}, region={FormatRegion(placement.Region)}");

            var animation = new DoubleAnimation(startX, endX, TimeSpan.FromSeconds(durationSeconds))
            {
                EasingFunction = null,
                FillBehavior = FillBehavior.Stop
            };
            animation.Completed += (sender, args) =>
            {
                if (sessionId != _displaySessionId)
                {
                    Debug.WriteLine($"[SlideAudience] displaySessionId mismatch ignored old={sessionId}, current={_displaySessionId}");
                    return;
                }

                Debug.WriteLine($"[SlideAudience] Flow item completed: {queued.Comment.Text}");
                RemoveActiveComment(active, "animationCompleted");
                TryDisplayNextComment(sessionId, forceImmediate: false);
            };
            item.BeginAnimation(Canvas.LeftProperty, animation);
        }

        private void CreateBubbleItem(QueuedComment queued, int sessionId)
        {
            var item = CreateAnimatedComment(queued.Comment);
            item.Opacity = 0;
            AnimationCanvas.Children.Add(item);
            item.UpdateLayout();

            var slotIndex = SelectBubbleSlot();
            var placement = SelectBubblePlacement(item, slotIndex, _currentWhitespaceRegions);
            Canvas.SetLeft(item, placement.Left);
            Canvas.SetTop(item, placement.Top);

            var fadeIn = GetBubbleFadeInSeconds();
            var hold = Math.Max(1, _settings.BubbleLifetimeSeconds);
            var fadeOut = GetBubbleFadeOutSeconds();
            var total = fadeIn + hold + fadeOut;
            var active = new ActiveComment
            {
                Comment = queued.Comment,
                CommentIndex = queued.CommentIndex,
                DisplayOrder = queued.DisplayOrder,
                LaneOrSlotIndex = slotIndex,
                Element = item,
                Mode = OverlayDisplayMode.Bubble,
                Bounds = new Rect(placement.Left, placement.Top, placement.Width, placement.Height)
            };
            _activeComments.Add(active);
            _shownCount++;

            Debug.WriteLine($"[SlideAudience] Bubble slot selected={slotIndex}");
            Debug.WriteLine($"[SlideAudience] Bubble slot index={slotIndex}");
            Debug.WriteLine($"[SlideAudience] Bubble x={placement.Left:F1}, y={placement.Top:F1}, fadeIn={fadeIn:F1}, hold={hold:F1}, fadeOut={fadeOut:F1}");
            Debug.WriteLine($"[SlideAudience] Bubble region selected index={placement.RegionIndex}, placementMode={placement.Mode}");
            if (!string.IsNullOrWhiteSpace(placement.FallbackAnchor))
            {
                Debug.WriteLine($"[SlideAudience] Bubble fallback anchor selected={placement.FallbackAnchor}");
            }

            Debug.WriteLine($"[SlideAudience] Bubble overlap retry count={placement.OverlapRetryCount}");
            Debug.WriteLine($"[SlideAudience] Bubble item created: {queued.Comment.Text}, selected region index={placement.RegionIndex}, placementMode={placement.Mode}, x={placement.Left:F1}, y={placement.Top:F1}, jitterX={placement.JitterX:F1}, jitterY={placement.JitterY:F1}, activeCommentsCount={_activeComments.Count}, queueRemaining={_commentQueue.Count}");
            Debug.WriteLine($"[SlideAudience] canvas children count={AnimationCanvas.Children.Count}");

            ScheduleLifecycleLog(sessionId, fadeIn, () => Debug.WriteLine($"[SlideAudience] Bubble fadeIn completed: {queued.Comment.Text}"));
            ScheduleLifecycleLog(sessionId, fadeIn + hold, () => Debug.WriteLine($"[SlideAudience] Bubble hold completed: {queued.Comment.Text}"));

            var storyboard = new Storyboard();
            var opacityAnimation = new DoubleAnimationUsingKeyFrames();
            opacityAnimation.KeyFrames.Add(new DiscreteDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.Zero)));
            opacityAnimation.KeyFrames.Add(new LinearDoubleKeyFrame(1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(fadeIn))));
            opacityAnimation.KeyFrames.Add(new LinearDoubleKeyFrame(1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(fadeIn + hold))));
            opacityAnimation.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(total))));
            Storyboard.SetTarget(opacityAnimation, item);
            Storyboard.SetTargetProperty(opacityAnimation, new PropertyPath(OpacityProperty));
            storyboard.Children.Add(opacityAnimation);
            storyboard.Completed += (sender, args) =>
            {
                if (sessionId != _displaySessionId)
                {
                    Debug.WriteLine($"[SlideAudience] displaySessionId mismatch ignored old={sessionId}, current={_displaySessionId}");
                    return;
                }

                Debug.WriteLine($"[SlideAudience] Bubble fadeOut completed: {queued.Comment.Text}");
                RemoveActiveComment(active, "fadeCompleted");
                TryDisplayNextComment(sessionId, forceImmediate: false);
            };
            storyboard.Begin();
        }

        private int SelectFlowLane()
        {
            var max = GetMaxSimultaneousComments();
            var activeLanes = _activeComments
                .Where(active => active.Mode == OverlayDisplayMode.Flow)
                .Select(active => active.LaneOrSlotIndex)
                .ToList();
            var available = Enumerable.Range(0, max)
                .Where(lane => !activeLanes.Contains(lane))
                .ToList();

            if (available.Count == 0)
            {
                return 0;
            }

            var preferred = available.Where(lane => lane != _lastFlowLane).ToList();
            var selected = (preferred.Count > 0 ? preferred : available)[_random.Next(preferred.Count > 0 ? preferred.Count : available.Count)];
            _lastFlowLane = selected;
            return selected;
        }

        private int SelectBubbleSlot()
        {
            var max = GetMaxSimultaneousComments();
            var activeSlots = _activeComments
                .Where(active => active.Mode == OverlayDisplayMode.Bubble)
                .Select(active => active.LaneOrSlotIndex)
                .ToList();
            var available = Enumerable.Range(0, max)
                .Where(slot => !activeSlots.Contains(slot))
                .ToList();
            return available.Count == 0 ? 0 : available[_random.Next(available.Count)];
        }

        private PlacementResult SelectFlowPlacement(int laneIndex, IReadOnlyList<WhitespaceRegion> whitespaceRegions)
        {
            var laneHeight = Math.Max(48, GetCommentFontSize() + 28);
            var selection = TrySelectWhitespaceRegion(whitespaceRegions, minWidth: 0.08, minHeight: 0.06);
            if (selection.Region != null)
            {
                var region = selection.Region;
                var minTop = Math.Max(0, region.Y * ActualHeight);
                var maxTop = Math.Max(minTop, (region.Y + region.Height) * ActualHeight - laneHeight);
                var top = AvoidNearbyTop(minTop + _random.NextDouble() * Math.Max(1, maxTop - minTop), laneHeight);
                _recentFlowTops.Add(top);
                TrimRecent(_recentFlowTops, 6);
                return new PlacementResult { Top = Math.Min(top, Math.Max(0, ActualHeight - laneHeight)), Mode = "whitespace", Region = region, RegionIndex = selection.Index };
            }

            Debug.WriteLine("[SlideAudience] placement fallback reason=no suitable flow whitespace region");
            var laneCount = GetMaxSimultaneousComments();
            var topMargin = Math.Max(40, ActualHeight * 0.12);
            var usableHeight = Math.Max(laneHeight, ActualHeight - topMargin * 2);
            var step = laneCount <= 1 ? 0 : usableHeight / laneCount;
            var fallbackTop = topMargin + laneIndex * step;
            return new PlacementResult { Top = Math.Min(fallbackTop, Math.Max(0, ActualHeight - laneHeight)), Mode = "fallback", RegionIndex = -1 };
        }

        private PlacementResult SelectBubblePlacement(
            FrameworkElement item,
            int slotIndex,
            IReadOnlyList<WhitespaceRegion> whitespaceRegions)
        {
            var itemWidth = EstimateCommentWidth(GetText(item), item);
            var itemHeight = Math.Max(48, item.ActualHeight);
            var selection = TrySelectWhitespaceRegion(
                whitespaceRegions,
                minWidth: Math.Max(0.12, (itemWidth + 40) / Math.Max(1, ActualWidth)),
                minHeight: Math.Max(0.08, (itemHeight + 20) / Math.Max(1, ActualHeight)));

            if (selection.Region != null)
            {
                var region = selection.Region;
                var regionLeft = region.X * ActualWidth;
                var regionTop = region.Y * ActualHeight;
                var regionWidth = region.Width * ActualWidth;
                var regionHeight = region.Height * ActualHeight;
                var baseLeft = regionLeft + _random.NextDouble() * Math.Max(1, regionWidth - itemWidth);
                var baseTop = regionTop + _random.NextDouble() * Math.Max(1, regionHeight - itemHeight);
                var jitterX = _random.NextDouble() * 80 - 40;
                var jitterY = _random.NextDouble() * 60 - 30;
                var result = ResolveBubbleOverlap(
                    baseLeft + jitterX,
                    baseTop + jitterY,
                    itemWidth,
                    itemHeight,
                    "whitespace",
                    selection.Index,
                    null,
                    jitterX,
                    jitterY);
                result.Region = region;
                RememberRegion(selection.Index);
                return result;
            }

            Debug.WriteLine("[SlideAudience] placement fallback reason=no suitable bubble whitespace region");
            return SelectBubbleFallbackPlacement(slotIndex, itemWidth, itemHeight);
        }

        private PlacementResult SelectBubbleFallbackPlacement(int slotIndex, double itemWidth, double itemHeight)
        {
            var anchors = _fallbackAnchors
                .Where(candidateAnchor => !string.Equals(candidateAnchor, _lastFallbackAnchor, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (anchors.Count == 0)
            {
                anchors = _fallbackAnchors.ToList();
            }

            var anchorName = anchors[_random.Next(anchors.Count)];
            _lastFallbackAnchor = anchorName;
            var anchorPoint = AnchorPoint(anchorName, itemWidth, itemHeight);
            var jitterX = _random.NextDouble() * 100 - 50;
            var jitterY = _random.NextDouble() * 80 - 40;
            var placement = ResolveBubbleOverlap(
                anchorPoint.X + jitterX,
                anchorPoint.Y + jitterY,
                itemWidth,
                itemHeight,
                "fallback",
                -1,
                anchorName,
                jitterX,
                jitterY);
            Debug.WriteLine($"[SlideAudience] Bubble fallback anchor selected={anchorName}");
            return placement;
        }

        private Point AnchorPoint(string anchorName, double itemWidth, double itemHeight)
        {
            var rightX = Math.Max(0, ActualWidth - itemWidth - Math.Max(40, _settings.MarginRight));
            var centerX = Math.Max(0, (ActualWidth - itemWidth) * 0.5);
            switch (anchorName)
            {
                case "rightTop":
                    return new Point(rightX, Math.Max(40, ActualHeight * 0.14));
                case "rightMiddle":
                    return new Point(rightX, Math.Max(40, (ActualHeight - itemHeight) * 0.48));
                case "rightBottom":
                    return new Point(rightX, Math.Max(40, ActualHeight - itemHeight - Math.Max(60, _settings.MarginBottom)));
                case "centerTop":
                    return new Point(centerX, Math.Max(40, ActualHeight * 0.18));
                default:
                    return new Point(centerX, Math.Max(40, ActualHeight - itemHeight - Math.Max(70, _settings.MarginBottom)));
            }
        }

        private PlacementResult ResolveBubbleOverlap(
            double left,
            double top,
            double width,
            double height,
            string mode,
            int regionIndex,
            string fallbackAnchor,
            double jitterX,
            double jitterY)
        {
            var bestLeft = Clamp(left, 0, Math.Max(0, ActualWidth - width));
            var bestTop = Clamp(top, 0, Math.Max(0, ActualHeight - height));
            var bestOverlap = OverlapScore(new Rect(bestLeft, bestTop, width, height));
            var retryCount = 0;

            for (var attempt = 0; attempt < 10; attempt++)
            {
                var candidateLeft = Clamp(left + (_random.NextDouble() * 160 - 80), 0, Math.Max(0, ActualWidth - width));
                var candidateTop = Clamp(top + (_random.NextDouble() * 120 - 60), 0, Math.Max(0, ActualHeight - height));
                var candidateRect = new Rect(candidateLeft, candidateTop, width, height);
                var overlap = OverlapScore(candidateRect);
                retryCount = attempt;
                if (overlap <= 0)
                {
                    bestLeft = candidateLeft;
                    bestTop = candidateTop;
                    bestOverlap = 0;
                    break;
                }

                if (overlap < bestOverlap)
                {
                    bestLeft = candidateLeft;
                    bestTop = candidateTop;
                    bestOverlap = overlap;
                }
            }

            Debug.WriteLine($"[SlideAudience] Bubble overlap retry count={retryCount}");
            return new PlacementResult
            {
                Left = bestLeft,
                Top = bestTop,
                Width = width,
                Height = height,
                Mode = mode,
                RegionIndex = regionIndex,
                FallbackAnchor = fallbackAnchor,
                JitterX = jitterX,
                JitterY = jitterY,
                OverlapRetryCount = retryCount
            };
        }

        private double OverlapScore(Rect rect)
        {
            return _activeComments
                .Where(active => active.Mode == OverlayDisplayMode.Bubble)
                .Select(active => Rect.Intersect(active.Bounds, rect))
                .Where(intersection => !intersection.IsEmpty)
                .Sum(intersection => intersection.Width * intersection.Height);
        }

        private RegionSelection TrySelectWhitespaceRegion(
            IReadOnlyList<WhitespaceRegion> whitespaceRegions,
            double minWidth,
            double minHeight)
        {
            if (!_settings.UseWhitespaceAwarePlacement)
            {
                Debug.WriteLine("[SlideAudience] placement fallback reason=whitespace-aware placement disabled");
                return new RegionSelection();
            }

            if (whitespaceRegions == null || whitespaceRegions.Count == 0)
            {
                Debug.WriteLine("[SlideAudience] placement fallback reason=no whitespace regions");
                return new RegionSelection();
            }

            if (_random.NextDouble() > GetWhitespacePlacementProbability())
            {
                Debug.WriteLine("[SlideAudience] placement fallback reason=probability chose legacy placement");
                return new RegionSelection();
            }

            var candidates = whitespaceRegions
                .Select((region, index) => new RegionSelection { Region = region, Index = index })
                .Where(selection => selection.Region.Width >= minWidth && selection.Region.Height >= minHeight)
                .OrderByDescending(selection => selection.Region.Score)
                .Take(Math.Min(10, whitespaceRegions.Count))
                .ToList();

            if (candidates.Count == 0)
            {
                Debug.WriteLine("[SlideAudience] placement fallback reason=whitespace region too small");
                return new RegionSelection();
            }

            var weighted = candidates
                .Select(selection => new
                {
                    Selection = selection,
                    Weight = Math.Max(0.05, selection.Region.Score) * (_recentRegionIndexes.Contains(selection.Index) ? 0.25 : 1.0)
                })
                .ToList();
            var total = weighted.Sum(item => item.Weight);
            var pick = _random.NextDouble() * total;
            foreach (var item in weighted)
            {
                pick -= item.Weight;
                if (pick <= 0)
                {
                    return item.Selection;
                }
            }

            return weighted.Last().Selection;
        }

        private void RememberRegion(int regionIndex)
        {
            if (regionIndex < 0)
            {
                return;
            }

            _recentRegionIndexes.Add(regionIndex);
            TrimRecent(_recentRegionIndexes, 5);
        }

        private double AvoidNearbyTop(double top, double laneHeight)
        {
            for (var attempt = 0; attempt < 6; attempt++)
            {
                if (!_recentFlowTops.Any(existing => Math.Abs(existing - top) < laneHeight * 0.85))
                {
                    return top;
                }

                top = Math.Max(0, Math.Min(ActualHeight - laneHeight, top + laneHeight));
            }

            return top;
        }

        private void RemoveActiveComment(ActiveComment active, string reason)
        {
            if (active == null)
            {
                return;
            }

            active.Element.BeginAnimation(Canvas.LeftProperty, null);
            active.Element.BeginAnimation(OpacityProperty, null);
            AnimationCanvas.Children.Remove(active.Element);
            _activeComments.Remove(active);
            Debug.WriteLine($"[SlideAudience] comment removed reason={reason}, commentIndex={active.CommentIndex}, displayOrder={active.DisplayOrder}, laneOrSlotIndex={active.LaneOrSlotIndex}");
            Debug.WriteLine($"[SlideAudience] {(active.Mode == OverlayDisplayMode.Flow ? "Flow" : "Bubble")} item removed: {active.Comment.Text}, reason={reason}");
            Debug.WriteLine($"[SlideAudience] active comments count={_activeComments.Count}");
            Debug.WriteLine($"[SlideAudience] canvas children count={AnimationCanvas.Children.Count}");
        }

        private void ScheduleLifecycleLog(int sessionId, double seconds, Action logAction)
        {
            var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(seconds) };
            timer.Tick += (sender, args) =>
            {
                timer.Stop();
                _lifecycleTimers.Remove(timer);
                if (sessionId == _displaySessionId)
                {
                    logAction();
                }
                else
                {
                    Debug.WriteLine($"[SlideAudience] displaySessionId mismatch ignored old={sessionId}, current={_displaySessionId}");
                }
            };
            _lifecycleTimers.Add(timer);
            timer.Start();
        }

        private void ResetPanelHideTimer()
        {
            StopPanelHideTimer();
            if (_settings.CommentLifetimeSeconds <= 0)
            {
                return;
            }

            _panelHideTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(_settings.CommentLifetimeSeconds)
            };
            _panelHideTimer.Tick += (sender, args) =>
            {
                StopPanelHideTimer();
                PanelBorder.Visibility = Visibility.Collapsed;
                Debug.WriteLine("[SlideAudience] Panel auto-hide");
            };
            _panelHideTimer.Start();
        }

        private void StopAllTimers()
        {
            StopPanelHideTimer();
            StopDisplayTimer();
            StopLifecycleTimers();
        }

        private void StopAllAnimations(string reason)
        {
            foreach (var active in _activeComments.ToList())
            {
                try
                {
                    active.Element.BeginAnimation(Canvas.LeftProperty, null);
                    active.Element.BeginAnimation(OpacityProperty, null);
                    Debug.WriteLine($"[SlideAudience] comment removed reason={reason}, text={active.Comment.Text}");
                    Debug.WriteLine($"[SlideAudience] {(active.Mode == OverlayDisplayMode.Flow ? "Flow" : "Bubble")} item removed: {active.Comment.Text}, reason={reason}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[SlideAudience] failed to stop animation reason={reason}, text={active.Comment.Text}");
                    Debug.WriteLine(ex.ToString());
                }
            }
        }

        private void StopPanelHideTimer()
        {
            if (_panelHideTimer == null)
            {
                return;
            }

            _panelHideTimer.Stop();
            _panelHideTimer = null;
        }

        private void StopDisplayTimer()
        {
            if (_displayTimer == null)
            {
                return;
            }

            _displayTimer.Stop();
            _displayTimer = null;
        }

        private void StopLifecycleTimers()
        {
            foreach (var timer in _lifecycleTimers.ToList())
            {
                timer.Stop();
            }

            _lifecycleTimers.Clear();
        }

        private int GetMaxSimultaneousComments()
        {
            return Math.Max(1, Math.Min(3, _settings.MaxSimultaneousComments));
        }

        private double GetNextCommentDelaySeconds()
        {
            var min = GetDisplayIntervalMinSeconds();
            var max = GetDisplayIntervalMaxSeconds();
            if (min > max)
            {
                var tmp = min;
                min = max;
                max = tmp;
            }

            return min + _random.NextDouble() * (max - min);
        }

        private double GetDisplayIntervalMinSeconds()
        {
            var value = _settings.CommentDisplayIntervalMinSeconds;
            if (value <= 0 && _settings.CommentDisplayIntervalSeconds > 0)
            {
                value = _settings.CommentDisplayIntervalSeconds;
            }

            return Clamp(value, 0.1, 30);
        }

        private double GetDisplayIntervalMaxSeconds()
        {
            var value = _settings.CommentDisplayIntervalMaxSeconds;
            if (value <= 0 && _settings.CommentDisplayIntervalSeconds > 0)
            {
                value = _settings.CommentDisplayIntervalSeconds;
            }

            return Clamp(value, 0.1, 30);
        }

        private double GetCommentFontSize()
        {
            return Clamp(_settings.CommentFontSize, 16, 48);
        }

        private double GetWhitespacePlacementProbability()
        {
            return Clamp(_settings.WhitespacePlacementProbability, 0, 1);
        }

        private double GetFlowSpeedPixelsPerSecond()
        {
            return Clamp(_settings.FlowSpeedPixelsPerSecond, 30, 200);
        }

        private double GetBubbleFadeInSeconds()
        {
            var value = _settings.BubbleFadeInSeconds > 0 ? _settings.BubbleFadeInSeconds : _settings.BubbleFadeSeconds;
            return Clamp(value, 0.1, 10);
        }

        private double GetBubbleFadeOutSeconds()
        {
            var value = _settings.BubbleFadeOutSeconds > 0 ? _settings.BubbleFadeOutSeconds : _settings.BubbleFadeSeconds;
            return Clamp(value, 0.1, 10);
        }

        private static Brush ParseBrushOrFallback(string color, string fallback)
        {
            try
            {
                var brush = (Brush)new BrushConverter().ConvertFromString(color);
                if (brush != null && brush.CanFreeze)
                {
                    brush.Freeze();
                }

                return brush;
            }
            catch
            {
                var fallbackBrush = (Brush)new BrushConverter().ConvertFromString(fallback);
                if (fallbackBrush != null && fallbackBrush.CanFreeze)
                {
                    fallbackBrush.Freeze();
                }

                return fallbackBrush;
            }
        }

        private double EstimateCommentWidth(string text, FrameworkElement item)
        {
            var actual = item != null && item.ActualWidth > 0 ? item.ActualWidth : 0;
            var estimated = Math.Max(220, (text ?? string.Empty).Length * GetCommentFontSize() * 1.1 + 48);
            return Math.Max(actual, estimated);
        }

        private static string GetText(FrameworkElement item)
        {
            var border = item as Border;
            var textBlock = border?.Child as TextBlock;
            return textBlock?.Text ?? string.Empty;
        }

        private static double Clamp(double value, double min, double max)
        {
            return Math.Max(min, Math.Min(max, value));
        }

        private static void TrimRecent<T>(List<T> values, int max)
        {
            while (values.Count > max)
            {
                values.RemoveAt(0);
            }
        }

        private static string FormatRegion(WhitespaceRegion region)
        {
            return region == null
                ? "(none)"
                : $"x={region.X:F3},y={region.Y:F3},w={region.Width:F3},h={region.Height:F3},score={region.Score:F3}";
        }

        private void OnSourceInitialized(object sender, EventArgs e)
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            var exStyle = GetWindowLongPtr(hwnd, GwlExstyle);
            SetWindowLongPtr(
                hwnd,
                GwlExstyle,
                new IntPtr(exStyle.ToInt64() | WsExTransparent | WsExToolWindow | WsExNoActivate));
        }

        private class QueuedComment
        {
            public AudienceComment Comment { get; set; }
            public int CommentIndex { get; set; }
            public int DisplayOrder { get; set; }
        }

        private class ActiveComment
        {
            public AudienceComment Comment { get; set; }
            public int CommentIndex { get; set; }
            public int DisplayOrder { get; set; }
            public int LaneOrSlotIndex { get; set; }
            public Border Element { get; set; }
            public OverlayDisplayMode Mode { get; set; }
            public Rect Bounds { get; set; }
        }

        private class PlacementResult
        {
            public double Left { get; set; }
            public double Top { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public string Mode { get; set; }
            public WhitespaceRegion Region { get; set; }
            public int RegionIndex { get; set; } = -1;
            public string FallbackAnchor { get; set; }
            public double JitterX { get; set; }
            public double JitterY { get; set; }
            public int OverlapRetryCount { get; set; }
        }

        private class RegionSelection
        {
            public WhitespaceRegion Region { get; set; }
            public int Index { get; set; } = -1;
        }

        private const int GwlExstyle = -20;
        private const long WsExTransparent = 0x00000020;
        private const long WsExToolWindow = 0x00000080;
        private const long WsExNoActivate = 0x08000000;

        private static IntPtr GetWindowLongPtr(IntPtr hwnd, int index)
        {
            return IntPtr.Size == 8
                ? GetWindowLongPtr64(hwnd, index)
                : new IntPtr(GetWindowLong32(hwnd, index));
        }

        private static IntPtr SetWindowLongPtr(IntPtr hwnd, int index, IntPtr newLong)
        {
            return IntPtr.Size == 8
                ? SetWindowLongPtr64(hwnd, index, newLong)
                : new IntPtr(SetWindowLong32(hwnd, index, newLong.ToInt32()));
        }

        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
    }
}
