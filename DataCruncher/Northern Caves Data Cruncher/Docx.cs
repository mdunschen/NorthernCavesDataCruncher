using System;
using System.Net;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Text;

namespace NCDataCruncher
{
  public static class Docx
  {
    public static string WidthFromCm(double cm)
    {
      return ((int)Math.Round(cm * 1440 / 2.54)).ToString();
    }

    public static TableCell FixedWidthCell(Paragraph paragraph, double widthCm)
    {
      var cellProps = new TableCellProperties(new TableCellWidth { Width = Docx.WidthFromCm(widthCm), Type = TableWidthUnitValues.Dxa });
      return new TableCell(cellProps, paragraph);
    }

    public static Style CreateStyle(string name, StyleValues styleType, string fontName = null, int? fontSize = null, bool bold = false, bool italic = false, string hexColour = null, Tabs tabs = null)
    {
      var style = new Style { Type = styleType, StyleId = name };
      var runProps = new StyleRunProperties();
      if (fontName != null)
        runProps.Append(new RunFonts { Ascii = fontName });
      if (fontSize.HasValue)
        runProps.Append(new FontSize { Val = fontSize.Value.ToString() });
      if (bold)
        runProps.Append(new Bold()); //?? this doesn't create bold text, it inverts the existing bold setting
      if (italic)
        runProps.Append(new Italic());
      style.Append(new Name { Val = name });
      if (hexColour != null)
        runProps.Append(new Color { Val = hexColour });
      //style.Append(new BasedOn { Val = "Heading1" });
      //style.Append(new NextParagraphStyle { Val = "Normal" });
      style.Append(runProps);

      if (tabs != null)
      {
        var paragraphProps = new StyleParagraphProperties(tabs);
        style.Append(paragraphProps);
      }
      return style;
    }

    //public static Style CreateTableStyle()
    //{
    //  var style = new Style();
    //  style.Type = StyleValues.Table
    //}

    public static Paragraph SimpleParagraph(string text, string style=null)
    {
      var paragraph = new Paragraph();
      if (style != null)
      {
        var paragraphProps = new ParagraphProperties();
        paragraphProps.ParagraphStyleId = new ParagraphStyleId() { Val = style };
        paragraph.Append(paragraphProps);
      }
      if (text != null)
        paragraph.Append(new Run(new Text(WebUtility.HtmlDecode(ConsolidateWhiteSpace(text.Trim())))));
      return paragraph;
    }

    public static Paragraph CreateParagraph(string styleName, params Run[] runs)
    {
      var paragraph = SimpleParagraph(null, styleName);
      foreach (Run run in runs)
        paragraph.Append(run);
      return paragraph;
    }

    public static Run StyledRun(string text, string style)
    {
      var run = new Run();
      RunProperties runProps = new RunProperties();
      runProps.RunStyle = new RunStyle { Val = style };
      run.Append(runProps);
      if (text != null)
//        run.Append(new Text(WebUtility.HtmlDecode(text)));
        run.Append(new Text(WebUtility.HtmlDecode(text)) { Space = SpaceProcessingModeValues.Preserve });
      return run;
    }

    public static void AddIndexMark(Paragraph paragraph, string text, string sortText)
    {
      if (sortText != null && sortText != "" && sortText != text)
        text = text + ";" + sortText;
      else if (text.StartsWith("The "))
        text = text.Substring(4) + ", The";

      paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
      paragraph.Append(new Run(new FieldCode { Space = SpaceProcessingModeValues.Preserve, Text = " XE \"" }));
//      paragraph.Append(new Run(new FieldCode { Text = text }) { RsidRunProperties = "00D03062" });
      paragraph.Append(new Run(new FieldCode { Text = text }));
      paragraph.Append(new Run(new FieldCode { Space = SpaceProcessingModeValues.Preserve, Text = "\" " }));
      paragraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }));
    }

    public static string ConsolidateWhiteSpace(string input)
    {
      bool whiteSpaceFound = false;
      StringBuilder output = new StringBuilder();
      foreach (char ch in input)
        if (ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r')
          whiteSpaceFound = true;
        else if (whiteSpaceFound)
        {
          whiteSpaceFound = false;
          output.Append(' ');
          output.Append(ch);
        }
        else
          output.Append(ch);
      if (whiteSpaceFound)
        output.Append(' ');
      return output.ToString();
    }
  }
}
