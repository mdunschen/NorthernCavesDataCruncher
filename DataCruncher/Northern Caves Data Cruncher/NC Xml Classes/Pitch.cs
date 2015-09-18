using System;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace NCDataCruncher
{
  public class Pitch: IXmlSerializable
  {
    public string Name { get; set; }
    public Rope Ladder { get; set; }
    public Rope Belay { get; set; }
    public Rope Lifeline { get; set; }
    public double? Height { get; set; }
    public Rope Rope { get; set; }
    public string SrtBelay { get; set; }

    public Pitch()
    {
      Ladder = new Rope();
      Belay = new Rope();
      Lifeline = new Rope();
      Rope = new Rope();
    }

    private double? ReadXmlDouble(XmlReader reader)
    {
      string str = reader.ReadElementString();
      double dbl;
      if (double.TryParse(str, out dbl))
        return dbl;
      return null;
    }

    public void ReadXml(XmlReader reader)
    {
      bool isEmpty = reader.IsEmptyElement;
      reader.ReadStartElement();
      if (isEmpty)
        return;
      while (reader.NodeType == XmlNodeType.Element)
      {
        if (reader.IsStartElement("name"))
          Name = reader.ReadElementString();
        else if (reader.IsStartElement("ladder"))
          Ladder.ReadXml(reader); // = Rope.FromXml(reader);
        else if (reader.IsStartElement("belay"))
          Belay.ReadXml(reader); // = Rope.FromXml(reader);
        else if (reader.IsStartElement("lifeline"))
          Lifeline.ReadXml(reader); // = Rope.FromXml(reader);
        else if (reader.IsStartElement("height"))
          Height = ReadXmlDouble(reader);
        else if (reader.IsStartElement("rope"))
          Rope.ReadXml(reader); // = Rope.FromXml(reader);
        else if (reader.IsStartElement("srtbelay"))
          SrtBelay = reader.ReadElementString();
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
  }
}
