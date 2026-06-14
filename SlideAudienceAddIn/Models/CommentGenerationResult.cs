using System;
using System.Collections.Generic;
using SlideAudienceAddIn.Services;

namespace SlideAudienceAddIn.Models
{
    public class CommentGenerationResult
    {
        public IReadOnlyList<AudienceComment> Comments { get; set; }

        public long LatencyMs { get; set; }

        public bool UsedApi { get; set; }

        public string ErrorMessage { get; set; }

        public string FallbackReason { get; set; }

        public string Model { get; set; }

        public int PromptLength { get; set; }

        public int ExtractedTextLength { get; set; }
    }

    public class SlideSnapshot
    {
        public string CacheKey { get; set; }

        public string PresentationHash { get; set; }

        public int SlideIndex { get; set; }

        public int SlideId { get; set; }

        public string ImagePath { get; set; }

        public string SlideText { get; set; }

        public long SlideExportTimeMs { get; set; }

        public long TextExtractTimeMs { get; set; }

        public DateTime CreatedAtUtc { get; set; }
    }

    public class SlidePreloadResult
    {
        public string CacheKey { get; set; }

        public int SlideIndex { get; set; }

        public int SlideId { get; set; }

        public string ImagePath { get; set; }

        public string SlideText { get; set; }

        public SlideWhitespaceAnalysisResult WhitespaceAnalysis { get; set; }

        public CommentGenerationResult CommentResult { get; set; }

        public bool UsedApi { get; set; }

        public bool Succeeded { get; set; }

        public string ErrorMessage { get; set; }

        public long SlideExportTimeMs { get; set; }

        public long TextExtractTimeMs { get; set; }

        public long WhitespaceAnalysisTimeMs { get; set; }

        public long CommentGenerationTimeMs { get; set; }

        public long TotalPreloadTimeMs { get; set; }

        public DateTime CreatedAtUtc { get; set; }
    }
}
