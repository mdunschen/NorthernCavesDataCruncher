using System.Xml.Serialization;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using DocumentFormat.OpenXml;

namespace NCDataCruncher
{
  public class Cave
  {
    [XmlElement("name")]
    public string Name { get; set; }
    [XmlElement("id")]
    public int Id { get; set; }
    [XmlElement("alias")]
    public string Alias { get; set; }
    [XmlElement("grid")]
    public string GridReference { get; set; }
    [XmlElement("elevation")]
    public double? Elevation { get; set; }
    [XmlElement("latitude")]
    public double? Latitude { get; set; }
    [XmlElement("longitude")]
    public double? Longitude { get; set; }
    [XmlElement("length")]
    public double? Length { get; set; }
    [XmlElement("length_note")]
    public string LengthNote { get; set; }
    [XmlElement("depth")]
    public double? Depth { get; set; }
    [XmlElement("depth_note")]
    public string DepthNote { get; set; }
    [XmlElement("grade")]
    public int? Grade { get; set; }
    [XmlElement("uppergrade")]
    public int? UpperGrade { get; set; }
    [XmlElement("warning")]
    public string Warning { get; set; }
    [XmlElement("permission")]
    public string Permission { get; set; }
    [XmlElement("rock")]
    public string Rock { get; set; }
    [XmlElement("exploration")]
    public string Exploration { get; set; }
    [XmlElement("description")]
    public Description Description { get; set; }
    [XmlElement("hydrology")]
    public string Hydrology { get; set; }
    [XmlElement("sortname")]
    public string SortName { get; set; }

    [XmlIgnore]
    //    public bool IsMajorCave { get { return Length.HasValue && Length.Value >= 100 || Depth.HasValue && Depth.Value >= 30; } }
    public bool IsMajorCave { get; set; }

    public void WriteDocx(Body body)
    {
      bool majorCave = IsMajorCave; // Length.HasValue && Length.Value >= 100 || Depth.HasValue && Depth.Value >= 30;
      Paragraph header = Docx.CreateParagraph(majorCave?StyleNames.CaveName:StyleNames.MinorCaveName, new Run(new Text(Docx.ConsolidateWhiteSpace(Name))));
      Docx.AddIndexMark(header, Name, CaveSortComparer.SortName(this));
      if (Grade.HasValue)
      {
        string gradeString = "Grade " + Grade.Value.ToString();
        if (UpperGrade.HasValue)
          gradeString = gradeString + "-" + UpperGrade.Value.ToString();
        if (majorCave)
          header.Append(new Run(new TabChar()), Docx.StyledRun(gradeString, StyleNames.Grade));
        else
          header.Append(new Run(new TabChar()), new Run(new Text(gradeString)));
      }
      body.Append(header);

      if (Alias != null)
      {
        Paragraph paragraph = Docx.SimpleParagraph(null, majorCave ? StyleNames.Alias : StyleNames.MinorAlias);
        string[] aliases = Alias.Split(',');
        for (int i = 0; i < aliases.Length; i++)
          aliases[i] = Docx.ConsolidateWhiteSpace(aliases[i]).Trim();
        paragraph.Append(new Run(new Text("(")));
        paragraph.Append(new Run(new Text(aliases[0])));
        Docx.AddIndexMark(paragraph, aliases[0], null);
        for (int i = 1; i < aliases.Length; i++)
        {
          paragraph.Append(new Run(new Text(", ") { Space = SpaceProcessingModeValues.Preserve }));
          paragraph.Append(new Run(new Text(aliases[i])));
          Docx.AddIndexMark(paragraph, aliases[i], null);
        }
        paragraph.Append(new Run(new Text(")")));
        body.Append(paragraph);
      }

      body.Append(LocationAndLengthParagraph(majorCave));

      if (LengthNote != null || DepthNote != null)
      {
        string lengthDepthNote = "";
        if (LengthNote != null && DepthNote != null)
          lengthDepthNote = LengthNote.Trim().TrimEnd('.') + ". " + DepthNote;
        else if (LengthNote != null)
          lengthDepthNote = LengthNote;
        else if (DepthNote != null)
          lengthDepthNote = DepthNote;
        body.Append(Docx.SimpleParagraph(lengthDepthNote, majorCave ? StyleNames.LengthDepthNote : StyleNames.MinorLengthDepthNote));
      }

      if (Exploration != null)
        body.Append(Docx.SimpleParagraph(Exploration, majorCave ? StyleNames.ExplorationHistory : StyleNames.MinorExplorationHistory));
      if (Warning != null)
        body.Append(Docx.SimpleParagraph("WARNING: " + Warning, majorCave ? StyleNames.Warning : StyleNames.MinorWarning));
      if (Description != null)
        Description.WriteDocx(body, majorCave, Rock);
      if (Hydrology != null)
        body.Append(Docx.SimpleParagraph("Hydrology: " + Hydrology, majorCave ? StyleNames.Hydrology : StyleNames.MinorHydrology)); // Sam request 1
      if (Permission != null)
        body.Append(Docx.SimpleParagraph("Permission: " + Permission, majorCave ? StyleNames.Permission : StyleNames.MinorPermission));
    }

    private Paragraph LocationAndLengthParagraph(bool majorCave)
    {
      var paragraph = Docx.SimpleParagraph(null, majorCave ? StyleNames.LocationAndLength : StyleNames.MinorLocationAndLength);

      if (GridReference != null)
      {
        var builder = new StringBuilder();
        foreach (char ch in GridReference)
          if (ch!=' ')
            builder.Append(ch);
        string gridReference = builder.ToString();
        if (gridReference.Length == 10)
          gridReference = string.Format("{0} {1} {2}", gridReference.Substring(0, 2), gridReference.Substring(2, 4), gridReference.Substring(6, 4));
        else if (gridReference.Length == 8)
          gridReference = string.Format("{0} {1} {2}", gridReference.Substring(0, 2), gridReference.Substring(2, 3), gridReference.Substring(5, 3));
        else
          gridReference = GridReference;
        paragraph.Append(new Run(new Text(gridReference)));
      }
      paragraph.Append(new Run(new TabChar()));
      if (Elevation.HasValue)
        paragraph.Append(new Run(new Text("Alt. " + Elevation.Value.ToString("F0") + "m")));
      paragraph.Append(new Run(new TabChar()));
      if (Length.HasValue && Depth.HasValue)
      {
        paragraph.Append(new Run(new Text("Length " + Length.Value.ToString("F0") + "m")));
        paragraph.Append(new Run(new TabChar()));
        paragraph.Append(new Run(new Text("Depth " + Depth.Value.ToString("F0") + "m")));
      }
      else if (Length.HasValue)
      {
        // Sam request 3
        //paragraph.Append(new Run(new TabChar())); 
        paragraph.Append(new Run(new Text("Length " + Length.Value.ToString("F0") + "m")));
      }
      else if (Depth.HasValue)
      {
        paragraph.Append(new Run(new TabChar()));
        paragraph.Append(new Run(new Text("Depth " + Depth.Value.ToString("F0") + "m")));
      }
      return paragraph;
    }
  }
}
