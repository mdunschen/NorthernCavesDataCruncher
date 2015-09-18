using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NCDataCruncher
{
  public class Para: IXmlSerializable, IDescriptionPart
  {
    public enum SpanType {Text, OffRoute, Sump, Sup};
    public class Span
    {
      public SpanType SpanType { get; set; }
      public string Text { get; set; }
    }

    public List<Span> Spans { get; set; }

    public Para()
    {
      Spans = new List<Span>();
    }

    public void ReadXml(XmlReader reader)
    {
      bool isEmpty = reader.IsEmptyElement;
      reader.ReadStartElement();
      if (isEmpty)
        return;
      while (reader.NodeType == XmlNodeType.Element || reader.NodeType == XmlNodeType.Text)
      {
        if (reader.NodeType == XmlNodeType.Text)
          Spans.Add(new Span { SpanType = SpanType.Text, Text = reader.ReadContentAsString() });
        else if (reader.IsStartElement("offroute"))
          Spans.Add(new Span { SpanType = SpanType.OffRoute, Text = reader.ReadElementString() });
        else if (reader.IsStartElement("sump"))
          Spans.Add(new Span { SpanType = SpanType.Sump, Text = reader.ReadElementString() });
        else if (reader.IsStartElement("sup"))
          throw new NotSupportedException("<sup> tag is not supported");
        else
          reader.ReadOuterXml();
      }
      reader.ReadEndElement();
    }

    public XmlSchema GetSchema()
    {
      throw new NotImplementedException();
    }

    public void WriteXml(XmlWriter writer)
    {
      throw new NotImplementedException();
    }

    //public void WriteDocx(Body body, bool majorCave)
    //{
    //  var paragraph = Docx.SimpleParagraph(null, majorCave? StyleNames.CaveDescription: StyleNames.MinorCaveDescription);
    //  bool firstSpan = true;
    //  bool prevSpanHasTrailingSpace = false;
    //  foreach (var span in Spans)
    //  {
    //    string text = ConsolidateWhiteSpace(WebUtility.HtmlDecode(span.Text));
    //    if (firstSpan || prevSpanHasTrailingSpace)
    //      text = text.TrimStart();
    //    prevSpanHasTrailingSpace = text.Length>0 && text[text.Length - 1] == ' ';
    //    firstSpan = false;
    //    if (span.SpanType == Para.SpanType.Text)
    //      //paragraph.Append(new Run(new Text(WebUtility.HtmlDecode(span.Text).Trim()) { Space = SpaceProcessingModeValues.Default }));
    //      paragraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })); 
    //    else if (span.SpanType == Para.SpanType.OffRoute)
    //      paragraph.Append(Docx.StyledRun(text, majorCave? StyleNames.OffRoute: StyleNames.MinorOffRoute));
    //    else if (span.SpanType == Para.SpanType.Sump)
    //      paragraph.Append(Docx.StyledRun(text, majorCave ? StyleNames.Sump : StyleNames.MinorSump));
    //  }
    //  body.Append(paragraph);
    //}

    public void WriteDocx(Body body, bool majorCave)
    {
        WriteDocx(body, majorCave, null);
    }

    public void WriteDocx(Body body, bool majorCave, string rock) // Sam request 2
    {
        var paragraph = Docx.SimpleParagraph(null, majorCave ? StyleNames.CaveDescription : StyleNames.MinorCaveDescription);
        bool firstSpan = true;
        if (rock != null)
        {
            paragraph.Append(Docx.StyledRun(rock + ".", StyleNames.Sump)); // Sam request 2
            firstSpan = false;
        }

      foreach (var span in Spans)
      {
        string text = Docx.ConsolidateWhiteSpace(WebUtility.HtmlDecode(span.Text)).Trim();
        if (!firstSpan && !(text.Length>0 && (".,;:)]".IndexOf(text[0]) != -1)))
          text = " " + text;
        firstSpan = false;
        if (span.SpanType == Para.SpanType.Text)
          paragraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        else if (span.SpanType == Para.SpanType.OffRoute)
          paragraph.Append(Docx.StyledRun(text, majorCave ? StyleNames.OffRoute : StyleNames.MinorOffRoute));
        else if (span.SpanType == Para.SpanType.Sump)
          paragraph.Append(Docx.StyledRun(text, majorCave ? StyleNames.Sump : StyleNames.MinorSump));
      }
      body.Append(paragraph);
    }
  }
}
