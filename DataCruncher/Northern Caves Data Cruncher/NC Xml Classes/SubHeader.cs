using System;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NCDataCruncher
{
  public class SubHeader: IDescriptionPart, IXmlSerializable
  {
    [XmlIgnore]
    public int Level { get; set; }
    [XmlElement("title")]
    public string Title { get; set; }
    [XmlElement("exploration")]
    public string Exploration { get; set; }
    [XmlElement("alias")]
    public string Alias { get; set; }
    [XmlElement("grade")]
    public int? Grade { get; set; }
    [XmlElement("upper_grade")]
    public int? UpperGrade { get; set; }
    [XmlElement("length")]
    public double? Length { get; set; }
    [XmlElement("warning")]
    public string Warning { get; set; }

    public SubHeader(int level)
    {
      Level = level;
    }

    public void ReadXml(XmlReader reader)
    {
      bool isEmpty = reader.IsEmptyElement;
      reader.ReadStartElement();
      if (isEmpty)
        return;
      while (reader.NodeType == XmlNodeType.Element || reader.NodeType == XmlNodeType.Text)
      {
        if (reader.IsStartElement("title"))
          Title = reader.ReadElementString();
        else if (reader.IsStartElement("exploration"))
          Exploration = reader.ReadElementString();
        else if (reader.IsStartElement("alias"))
          Alias = reader.ReadElementString();
        else if (reader.IsStartElement("grade"))
          Grade = reader.ReadElementContentAsInt();
        else if (reader.IsStartElement("upper_grade"))
          UpperGrade = reader.ReadElementContentAsInt();
        else if (reader.IsStartElement("length"))
          Length = reader.ReadElementContentAsDouble();
        else if (reader.IsStartElement("warning"))
          Warning = reader.ReadElementString();
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

    public void WriteDocx(Body body, bool majorCave)
    {
      string styleName = Level==1?StyleNames.SubHeader:StyleNames.SubHeader2;
      Paragraph header = Docx.CreateParagraph(styleName, new Run(new Text(Docx.ConsolidateWhiteSpace(Title))));
      if (Grade.HasValue)
      {
        string gradeString = "Grade " + Grade.Value.ToString();
        if (UpperGrade.HasValue)
          gradeString = gradeString + "-" + UpperGrade.Value.ToString();
        header.Append(new Run(new TabChar()), Docx.StyledRun(gradeString, styleName));
      }
      body.Append(header);

      if (Alias!=null)
        body.Append(Docx.SimpleParagraph("(" + Alias + ")", StyleNames.Alias));

      if (Length!=null)
      {
        string lengthText = "Length " + Length.Value.ToString("F0");
        Paragraph para = Docx.CreateParagraph(StyleNames.LocationAndLength, new Run(new TabChar()), new Run(new TabChar()), new Run(new TabChar()), new Run(new Text(lengthText)));
        body.Append(para);
      }
          
      if (Exploration != null)
        body.Append(Docx.SimpleParagraph(Exploration, StyleNames.ExplorationHistory));      
      
      if (Warning != null)
        body.Append(Docx.SimpleParagraph("WARNING: " + Warning, StyleNames.Warning));
    }
  }
}
