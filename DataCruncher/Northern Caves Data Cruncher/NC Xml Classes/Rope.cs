using System;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace NCDataCruncher
{
  public class Rope: IXmlSerializable
  {
    public string Text { get; set; }
    public double? Length { get; set; }
    public int? Span { get; set; }

    public void ReadXml(XmlReader reader)
    {
      reader.MoveToAttribute("span");
      if (reader.NodeType == XmlNodeType.Attribute)
        Span = (int)reader.ReadContentAsDouble();
      reader.MoveToElement();
      Text = reader.ReadElementContentAsString();
      double length;
      if (double.TryParse(Text, out length))
        Length = length;
    }

    public XmlSchema GetSchema()
    {
      throw new NotImplementedException();
    }

    public void WriteXml(XmlWriter writer)
    {
      throw new NotImplementedException();
    }

    //public static Rope FromXml(XmlReader reader)
    //{
    //  var rope = new Rope();
    //  rope.ReadXml(reader);
    //  return rope;
    //}

    public override string ToString()
    {
      if (Length.HasValue)
        return Length.Value.ToString() + "m";
      else
        return Text;
    }

    //private double? ReadXmlDouble(XmlReader reader)
    //{
    //  string str = reader.ReadElementString();
    //  double dbl;
    //  if (double.TryParse(str, out dbl))
    //    return dbl;
    //  return null;
    //}
  }
}
