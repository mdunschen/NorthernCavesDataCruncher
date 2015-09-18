using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Xml;
using System.Xml.Schema;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NCDataCruncher
{
  public class Description: IXmlSerializable
  {
    [XmlElement("p")]
    public List<IDescriptionPart> Parts { get; set; }

    public Description()
    {
      Parts = new List<IDescriptionPart>();
    }

    public void ReadXml(XmlReader reader)
    {
      bool isEmpty = reader.IsEmptyElement;
      reader.ReadStartElement();
      if (isEmpty)
        return;
      while (reader.NodeType == XmlNodeType.Element)
      {
        if (reader.IsStartElement("p"))
        {
          var para = new Para();
          para.ReadXml(reader);
          Parts.Add(para);
        }
        else if (reader.IsStartElement("subheader"))
        {
          var subHeader = new SubHeader(1);
          subHeader.ReadXml(reader);
          Parts.Add(subHeader);
        }
        else if (reader.IsStartElement("subheader2"))
        {
          var subHeader = new SubHeader(2);
          subHeader.ReadXml(reader);
          Parts.Add(subHeader);
        }
        else if (reader.IsStartElement("pitches"))
        {
          var pitches = new Pitches();
          pitches.ReadXml(reader);
          Parts.Add(pitches);
        }
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

    public void WriteDocx(Body body, bool majorCave, string rock)
    {
      bool prevPitches = false;
      foreach (var part in Parts)
      {
        if (part is Pitches)
        {
          (part as Pitches).WriteDocx(body, majorCave, prevPitches);
          prevPitches = true;
        }
        else if (part is Para)
        {
            (part as Para).WriteDocx(body, majorCave, rock);
            rock = null;
        }
        else
        {
          part.WriteDocx(body, majorCave);
          prevPitches = false;
        }

      }
    }
  }
}
