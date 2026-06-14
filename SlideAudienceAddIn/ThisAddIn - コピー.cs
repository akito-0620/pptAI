using System;
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideAudienceAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Debug.WriteLine("[SlideAudience] Add-in startup");

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
            }
            catch
            {
                // Ignore shutdown cleanup errors.
            }
        }

        private void Application_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            Debug.WriteLine("[SlideAudience] Slide show begin");
            LogCurrentSlide(Wn, "begin");
        }

        private void Application_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            Debug.WriteLine("[SlideAudience] Slide show next slide");
            LogCurrentSlide(Wn, "next");
        }

        private void Application_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            Debug.WriteLine("[SlideAudience] Slide show end");

            try
            {
                string name = Pres != null ? Pres.Name : "(unknown presentation)";
                Debug.WriteLine($"[SlideAudience] Presentation ended: {name}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] Error on slideshow end: {ex.Message}");
            }
        }

        private void LogCurrentSlide(PowerPoint.SlideShowWindow Wn, string eventName)
        {
            try
            {
                if (Wn == null || Wn.View == null || Wn.View.Slide == null)
                {
                    Debug.WriteLine($"[SlideAudience] {eventName}: slide window/view/slide is null");
                    return;
                }

                PowerPoint.Slide slide = Wn.View.Slide;

                int slideIndex = slide.SlideIndex;
                int slideId = slide.SlideID;
                string slideName = slide.Name;

                Debug.WriteLine(
                    $"[SlideAudience] {eventName}: index={slideIndex}, id={slideId}, name={slideName}"
                );
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SlideAudience] Error while logging current slide: {ex.Message}");
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