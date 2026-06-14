using System;
using System.Diagnostics;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideAudienceAddIn.Services
{
    public class SlideExporter
    {
        public string ExportSlideAsPng(PowerPoint.Slide slide, int width = 1280, int height = 720)
        {
            var dir = Path.Combine(Path.GetTempPath(), "SlideAudience", "exports");
            Directory.CreateDirectory(dir);

            var filePath = Path.Combine(
                dir,
                $"slide_{slide.SlideID}_{DateTime.Now:yyyyMMdd_HHmmss_fff}.png");

            Debug.WriteLine($"[SlideAudience] Slide.Export start slideId={slide.SlideID}, slideIndex={slide.SlideIndex}, path={filePath}, width={width}, height={height}");
            slide.Export(filePath, "PNG", width, height);
            Debug.WriteLine($"[SlideAudience] Slide.Export success=True slideId={slide.SlideID}, slideIndex={slide.SlideIndex}, path={filePath}, width={width}, height={height}");
            return filePath;
        }
    }
}
