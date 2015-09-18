using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NCDataCruncher
{
  public class Pitches : IDescriptionPart, IXmlSerializable
  {
    public class PitchSequence
    {
      public string Title { get; set; }
      public List<Pitch> Pitches { get; set; }

      public PitchSequence()
      {
        Pitches = new List<Pitch>();
      }
    }

    public List<PitchSequence> PitchSequences { get; set; }

    public Pitches()
    {
      PitchSequences = new List<PitchSequence>();
    }

    public void ReadXml(XmlReader reader)
    {
      int sequenceCount = 0;
      bool isEmpty = reader.IsEmptyElement;
      reader.ReadStartElement();
      if (isEmpty)
        return;
      while (reader.NodeType == XmlNodeType.Element)
      {
        if (sequenceCount == 0 || reader.IsStartElement("title"))
        {
          PitchSequences.Add(new PitchSequence());
          sequenceCount++;
        }
        if (reader.IsStartElement("title"))
          PitchSequences[sequenceCount - 1].Title = reader.ReadElementString();
        else if (reader.IsStartElement("pitch"))
        {
          var pitch = new Pitch();
          pitch.ReadXml(reader);
          PitchSequences[sequenceCount - 1].Pitches.Add(pitch);
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

    public void WriteDocx(Body body, bool majorCave)
    {
      WriteDocx(body, majorCave, false);
    }

    public void WriteDocx(Body body, bool majorCave, bool suppressHeader)
    {
      if (PitchSequences.Count == 1 && PitchSequences[0].Pitches.Count == 1 && (PitchSequences[0].Title == null || PitchSequences[0].Title == ""))
        WriteSimpleTackleList(PitchSequences[0].Pitches[0], body, majorCave);
      else
      {
        // Sam request 4
        //if (!suppressHeader)
        //  body.Append(Docx.SimpleParagraph("TTTTackle", StyleNames.PitchTableHeader)); 
        foreach (PitchSequence pitchSequence in PitchSequences)
          WritePitchSequenceDocx(pitchSequence, body);
      }
    }

    private void WriteSimpleTackleList(Pitch pitch, Body body, bool majorCave)
    {
      var builder = new StringBuilder();
      if (pitch.Name != null && pitch.Name != "" && pitch.Name != "Entrance" && pitch.Name != "1st" && pitch.Name != "First")
        builder.Append(" (" + pitch.Name + ") ");
     
      bool ladder = false; 
      if (pitch.Ladder != null && pitch.Ladder.Text != null && pitch.Ladder.Text != "")
      {
        builder.Append(pitch.Ladder.ToString());
        if (pitch.Ladder.Length.HasValue)
          builder.Append(" ladder");
        ladder = true;
      }
      if (pitch.Belay != null && pitch.Belay.Text != null && pitch.Belay.Text != "")
      {
        if (ladder)
           builder.Append("; ");
        string belayString = AnchorString(pitch.Belay.ToString()) + " belay";
        builder.Append(belayString);
        ladder = true;
      }
      if (pitch.Lifeline != null && pitch.Lifeline.Text != null && pitch.Lifeline.Text != "")
      {
        builder.Append((ladder ? "; " : "") + pitch.Lifeline.ToString());
        if (pitch.Lifeline.Length.HasValue)
          builder.Append(" lifeline");
        ladder = true;
      }

      StringBuilder builder2 = new StringBuilder();
      bool rope = false;
      if (pitch.Rope != null && pitch.Rope.Text != null && pitch.Rope.Text != "")
      {
        builder2.Append(pitch.Rope.ToString());
        if (pitch.Rope.Length.HasValue)
          builder2.Append(" rope");
        rope = true;
      }
      if (pitch.SrtBelay != null && pitch.SrtBelay != "")
      {
        if (rope)
          builder2.Append("; ");
        string anchorString = AnchorString(pitch.SrtBelay);
        if (anchorString=="Bar" || anchorString=="Stake")
          builder2.Append(anchorString + " belay");
        else
          builder2.Append(anchorString + " anchors");
        rope = true;
      }
      if (ladder & rope)
        builder.Append(" or ");
      if (rope)
        builder.Append(builder2.ToString());

      Run tackleRun = Docx.StyledRun(builder.ToString(), majorCave ? StyleNames.Tackle : StyleNames.MinorTackle);
      Paragraph paragraph = Docx.CreateParagraph(majorCave ? StyleNames.TackleHeader : StyleNames.MinorTackleHeader, new Run(new Text("Tackle: ") { Space = SpaceProcessingModeValues.Preserve }));
      paragraph.Append(tackleRun);
      body.Append(paragraph);
    }

    private string AnchorString(string code)
    {
    /*
      if (code == "F")
        return "fixed bolt";
      else if (code == "P")
        return "permanent bolt";
      else if (code == "S")
        return "threaded sleeve";
      else if (code == "N")
        return "natural";
      else if (code == "T")
        return "through bolt";
      else if (code == "B")
        return "bar";
      else if (code == "K")
        return "stake";
      else */ 
    // Sam request 5
        return code;
    }

    private string SrtAnchorString(string code)
    {
      if (code == "F")
        return "Fixed";
      else if (code == "P")
        return "Fixed";
      else if (code == "S")
        return "Spits";
      else if (code == "N")
        return "Natural";
      else if (code == "T")
        return "Through bolts";
      else if (code == "B")
        return "Bar";
      else if (code == "K")
        return "Stake";
      else
        return code;
    }

    private void WritePitchSequenceDocx(PitchSequence pitchSequence, Body body)
    {
      if (pitchSequence.Title != null)
        body.Append(Docx.SimpleParagraph(pitchSequence.Title, StyleNames.PitchTableHeader));
      if (pitchSequence.Pitches.Count > 0)
      {
        var table = new Table();

        TableProperties tableProps = new TableProperties(new TableStyle { Val = "ncTackleTable" });
        table.Append(tableProps);

        bool ladder = false;
        bool srt = false;
        foreach (Pitch pitch in pitchSequence.Pitches)
        {
          if (pitch.Ladder != null && pitch.Ladder.Text != null && pitch.Ladder.Text != "" ||
            pitch.Lifeline != null && pitch.Lifeline.Text != null && pitch.Lifeline.Text != "" || 
            pitch.Belay != null && pitch.Belay.Text != null && pitch.Belay.Text != "")
            ladder = true;
          if (pitch.Rope != null && pitch.Rope.Text != null && pitch.Rope.Text != "" || 
            pitch.SrtBelay != null && pitch.SrtBelay != null && pitch.SrtBelay != "")
            srt = true;
        }

        const double pageWidth = 12.3; //??
        const double NarrowColWidth = 1.5;
        const double WideColWidth = 2;
        double firstColWidth;
        if (ladder && srt)
          firstColWidth = pageWidth - 5 * NarrowColWidth;
        else if (ladder)
          firstColWidth = pageWidth - 3 * WideColWidth;
        else
          firstColWidth = pageWidth - 2 * WideColWidth;
        double otherColWidth = ladder && srt ? NarrowColWidth : WideColWidth;
        var firstRow = new TableRow();
        firstRow.Append(Docx.FixedWidthCell(Docx.SimpleParagraph("Pitch"), firstColWidth));
        if (ladder)
          firstRow.Append(
            Docx.FixedWidthCell(Docx.SimpleParagraph("Ladder"), otherColWidth),
            Docx.FixedWidthCell(Docx.SimpleParagraph("Belay"), otherColWidth),
            Docx.FixedWidthCell(Docx.SimpleParagraph("Lifeline"), otherColWidth));
        if (srt)
          firstRow.Append(
            Docx.FixedWidthCell(Docx.SimpleParagraph("Rope"), otherColWidth),
            Docx.FixedWidthCell(Docx.SimpleParagraph("Anchors"), otherColWidth));
        table.Append(firstRow);


        foreach (Pitch pitch in pitchSequence.Pitches)
        {
          var row = new TableRow();
          string pitchName = pitch.Name + (pitch.Height != null ? string.Format(" ({0}m)", pitch.Height) : "");
          row.Append(new TableCell(Docx.SimpleParagraph(pitchName)));
          if (ladder)
          {
            //double ladderLength;
            //double belayLength;
            //string ladderString = double.TryParse(pitch.Ladder, out ladderLength) ? ladderLength.ToString() + "m" : pitch.Ladder;
            //string belayString = double.TryParse(pitch.Belay, out belayLength) ? belayLength.ToString() + "m" : pitch.Belay;
            //string lifelineString = pitch.Lifeline.HasValue ? pitch.Lifeline.ToString() + "m" : "";
            row.Append(
              new TableCell(Docx.SimpleParagraph(pitch.Ladder.ToString())),
              new TableCell(Docx.SimpleParagraph(pitch.Belay.ToString())),
              new TableCell(Docx.SimpleParagraph(pitch.Lifeline.ToString())));
          }
          if (srt)
          {
            //string rope = pitch.Rope != null && pitch.Rope.Text != "") //?? span
            row.Append(
              new TableCell(Docx.SimpleParagraph(pitch.Rope.ToString())),
              new TableCell(Docx.SimpleParagraph(SrtAnchorString(pitch.SrtBelay))));
          }
          table.Append(row);
        }
        body.Append(table);
      }
    }
  }
}
