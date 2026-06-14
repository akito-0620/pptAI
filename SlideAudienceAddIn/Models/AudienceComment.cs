namespace SlideAudienceAddIn.Models
{
    public class AudienceComment
    {
        public string Type { get; set; }

        public string Persona { get; set; }

        public string Text { get; set; }

        public double? Confidence { get; set; }

        public static AudienceComment Create(string type, string text, double? confidence = null, string persona = null)
        {
            return new AudienceComment
            {
                Type = type,
                Persona = persona,
                Text = text,
                Confidence = confidence
            };
        }
    }
}
