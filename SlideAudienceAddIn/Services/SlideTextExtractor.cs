using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideAudienceAddIn.Services
{
    public class SlideTextExtractor
    {
        public string ExtractText(PowerPoint.Slide slide)
        {
            var sb = new StringBuilder();

            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                ExtractShapeText(shape, sb);
            }

            return sb.ToString().Trim();
        }

        private void ExtractShapeText(PowerPoint.Shape shape, StringBuilder sb)
        {
            try
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue &&
                    shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    AppendIfPresent(shape.TextFrame.TextRange.Text, sb);
                }

                if (shape.HasTable == MsoTriState.msoTrue)
                {
                    ExtractTableText(shape, sb);
                }

                if (shape.Type == MsoShapeType.msoGroup)
                {
                    foreach (PowerPoint.Shape child in shape.GroupItems)
                    {
                        ExtractShapeText(child, sb);
                    }
                }
            }
            catch
            {
                // Some PowerPoint shapes throw when inspected through interop. Continue with the rest.
            }
        }

        private static void ExtractTableText(PowerPoint.Shape shape, StringBuilder sb)
        {
            try
            {
                var table = shape.Table;
                for (var row = 1; row <= table.Rows.Count; row++)
                {
                    for (var col = 1; col <= table.Columns.Count; col++)
                    {
                        var cellShape = table.Cell(row, col).Shape;
                        if (cellShape.HasTextFrame == MsoTriState.msoTrue &&
                            cellShape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            AppendIfPresent(cellShape.TextFrame.TextRange.Text, sb);
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private static void AppendIfPresent(string text, StringBuilder sb)
        {
            if (!string.IsNullOrWhiteSpace(text))
            {
                sb.AppendLine(text.Trim());
            }
        }
    }
}
