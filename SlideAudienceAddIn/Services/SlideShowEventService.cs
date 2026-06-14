using System;
using System.Threading;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using SlideAudienceAddIn.Models;
using SlideAudienceAddIn.Overlay;

namespace SlideAudienceAddIn.Services
{
    public class SlideShowEventService
    {
        private readonly PowerPoint.Application _application;
        private readonly OverlayController _overlayController;
        private readonly CommentGenerationService _commentGenerationService;
        private readonly SlideExporter _slideExporter;
        private readonly SlideTextExtractor _slideTextExtractor;
        private readonly CommentCache _commentCache;
        private readonly ExperimentLogger _logger;
        private readonly AppSettings _settings;
        private CancellationTokenSource _currentCts;
        private int _currentSlideId;
        private bool _registered;

        public SlideShowEventService(
            PowerPoint.Application application,
            OverlayController overlayController,
            CommentGenerationService commentGenerationService,
            SlideExporter slideExporter,
            SlideTextExtractor slideTextExtractor,
            CommentCache commentCache,
            ExperimentLogger logger,
            AppSettings settings)
        {
            _application = application;
            _overlayController = overlayController;
            _commentGenerationService = commentGenerationService;
            _slideExporter = slideExporter;
            _slideTextExtractor = slideTextExtractor;
            _commentCache = commentCache;
            _logger = logger;
            _settings = settings;
        }

        public void RegisterEvents()
        {
            if (_registered)
            {
                return;
            }

            _application.SlideShowBegin += OnSlideShowBegin;
            _application.SlideShowNextSlide += OnSlideShowNextSlide;
            _application.SlideShowEnd += OnSlideShowEnd;
            _registered = true;
        }

        public void UnregisterEvents()
        {
            if (!_registered)
            {
                return;
            }

            _application.SlideShowBegin -= OnSlideShowBegin;
            _application.SlideShowNextSlide -= OnSlideShowNextSlide;
            _application.SlideShowEnd -= OnSlideShowEnd;
            _currentCts?.Cancel();
            _registered = false;
        }

        public void ClearCache()
        {
            _commentCache.Clear();
        }

        public Task GenerateForCurrentSlideAsync()
        {
            var window = TryGetActiveSlideShowWindow();
            if (window == null)
            {
                return Task.CompletedTask;
            }

            return ProcessSlideAsync(window);
        }

        public async Task<int> PregenerateForActivePresentationAsync(CancellationToken cancellationToken)
        {
            PowerPoint.Presentation presentation;
            try
            {
                presentation = _application.ActivePresentation;
            }
            catch
            {
                return 0;
            }

            var generatedCount = 0;
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var slideId = slide.SlideID;
                if (_commentCache.TryGet(slideId, out _))
                {
                    continue;
                }

                var slideIndex = slide.SlideIndex;
                var slideImagePath = _slideExporter.ExportSlideAsPng(slide);
                var extractedText = _slideTextExtractor.ExtractText(slide);
                _logger.LogSlideAnalyzed(slideId, slideIndex, slideImagePath, extractedText);

                var result = await _commentGenerationService.GenerateForSlideAsync(
                    slideId,
                    slideIndex,
                    slideImagePath,
                    extractedText,
                    cancellationToken);

                _commentCache.Set(slideId, result.Comments);
                _logger.LogCommentsGenerated(slideId, slideIndex, result);
                generatedCount++;
            }

            return generatedCount;
        }

        private void OnSlideShowBegin(PowerPoint.SlideShowWindow window)
        {
            if (_settings.Enabled)
            {
                _ = ProcessSlideAsync(window);
            }
        }

        private void OnSlideShowNextSlide(PowerPoint.SlideShowWindow window)
        {
            if (_settings.Enabled)
            {
                _ = ProcessSlideAsync(window);
            }
        }

        private void OnSlideShowEnd(PowerPoint.Presentation presentation)
        {
            _currentCts?.Cancel();
            _overlayController.Hide();
        }

        private async Task ProcessSlideAsync(PowerPoint.SlideShowWindow window)
        {
            _currentCts?.Cancel();
            _currentCts = new CancellationTokenSource();
            var token = _currentCts.Token;

            PowerPoint.Slide slide;
            try
            {
                slide = window.View.Slide;
            }
            catch
            {
                return;
            }

            var slideId = slide.SlideID;
            var slideIndex = slide.SlideIndex;
            _currentSlideId = slideId;

            try
            {
                _overlayController.AttachToSlideShowWindow(window);
                _logger.LogSlideChanged(slideId, slideIndex, TryGetPresentationPath(slide));

                if (_commentCache.TryGet(slideId, out var cachedComments))
                {
                    _overlayController.ShowComments(cachedComments);
                    _logger.LogCommentsShown(slideId, slideIndex, cachedComments);
                    return;
                }

                _overlayController.ShowLoading(slideIndex);

                var slideImagePath = _slideExporter.ExportSlideAsPng(slide);
                var extractedText = _slideTextExtractor.ExtractText(slide);
                _logger.LogSlideAnalyzed(slideId, slideIndex, slideImagePath, extractedText);
                var result = await _commentGenerationService.GenerateForSlideAsync(
                    slideId,
                    slideIndex,
                    slideImagePath,
                    extractedText,
                    token);

                if (token.IsCancellationRequested || _currentSlideId != slideId)
                {
                    return;
                }

                _commentCache.Set(slideId, result.Comments);
                _logger.LogCommentsGenerated(slideId, slideIndex, result);
                _overlayController.ShowComments(result.Comments);
                _logger.LogCommentsShown(slideId, slideIndex, result.Comments);
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                if (!token.IsCancellationRequested && _currentSlideId == slideId)
                {
                    _overlayController.ShowError("コメント生成に失敗しました");
                }

                _logger.LogError(slideId, slideIndex, ex);
            }
        }

        private PowerPoint.SlideShowWindow TryGetActiveSlideShowWindow()
        {
            try
            {
                return _application.SlideShowWindows.Count > 0
                    ? _application.SlideShowWindows[1]
                    : null;
            }
            catch
            {
                return null;
            }
        }

        private static string TryGetPresentationPath(PowerPoint.Slide slide)
        {
            try
            {
                var presentation = slide.Parent as PowerPoint.Presentation;
                return presentation?.FullName;
            }
            catch
            {
                return null;
            }
        }
    }
}
