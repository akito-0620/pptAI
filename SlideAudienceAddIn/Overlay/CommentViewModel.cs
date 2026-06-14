namespace SlideAudienceAddIn.Overlay
{
    public class CommentViewModel
    {
        public string Type { get; set; }

        public string Text { get; set; }

        public string AccentBrushKey
        {
            get
            {
                switch (Type)
                {
                    case "understanding":
                        return "UnderstandingBrush";
                    case "interest":
                        return "InterestBrush";
                    case "question":
                        return "QuestionBrush";
                    default:
                        return "DefaultAccentBrush";
                }
            }
        }
    }
}
