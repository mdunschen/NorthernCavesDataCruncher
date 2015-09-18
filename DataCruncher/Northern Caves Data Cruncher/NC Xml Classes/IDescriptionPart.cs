using DocumentFormat.OpenXml.Wordprocessing;

namespace NCDataCruncher
{
  public interface IDescriptionPart
  {
    void WriteDocx(Body body, bool MajorCave);
  }
}
