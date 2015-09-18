using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;

namespace NCDataCruncher
{
  [XmlRoot("area")]
  public class Area
  {
    [XmlElement("name")]
    public string Name { get; set; }
    [XmlElement("id")]
    public int Id { get; set; }
    [XmlElement("cave")]
    public List<Cave> Caves { get; set; }

    public Area()
    {
      Caves = new List<Cave>();
    }

    public void Save(string filePath)
    {
      XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
      ns.Add("", "");
      XmlSerializer serializer = new XmlSerializer(typeof(Area));
      using (Stream stream = File.Create(filePath))
        serializer.Serialize(stream, this, ns);
    }

    public static Area Load(string filePath)
    {
      Area area = null;
      var settings = new XmlReaderSettings();
      settings.IgnoreComments = true;
      settings.IgnoreProcessingInstructions = true;
      settings.IgnoreWhitespace = true;
      XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
      ns.Add("", "");
      using (var reader = XmlReader.Create(filePath, settings))
      {
        if (reader.IsStartElement("area",""))
        {
          string defaultNamespace = reader["xmlns"];
          XmlSerializer serializer = new XmlSerializer(typeof(Area),defaultNamespace);
          area = (Area)serializer.Deserialize(reader);
        }
      }
      return area;
    }

    public void WriteDocx(Body body, Func<string, string, bool> majorCave)
    {
      List<Cave> majorCaves = new List<Cave>();
      List<Cave> minorCaves = new List<Cave>();
      body.Append(Docx.SimpleParagraph(Name, StyleNames.AreaHeader));

      foreach (Cave cave in Caves)
      {
        cave.IsMajorCave = majorCave(Name, cave.Name);
        if (cave.IsMajorCave)
          majorCaves.Add(cave);
        else
          minorCaves.Add(cave);
      }

      majorCaves.Sort(new CaveSortComparer());
      minorCaves.Sort(new CaveSortComparer());
      foreach (Cave cave in majorCaves)
        cave.WriteDocx(body);
      if (minorCaves.Count > 0)
      {
        body.Append(Docx.SimpleParagraph("Minor Caves", StyleNames.MinorCavesHeader));
        foreach (Cave cave in minorCaves)
          cave.WriteDocx(body);
      }
    }
  }
}
