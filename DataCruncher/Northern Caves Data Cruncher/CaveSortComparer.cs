using System.Collections.Generic;

namespace NCDataCruncher
{
  public class CaveSortComparer : IComparer<Cave>
  {
    public static string SortName(Cave cave)
    {
      string name;
      if (cave.SortName != null && cave.SortName != "")
        name = cave.SortName;
      else
        name = cave.Name;
      if (name.StartsWith("The"))
        return name.Substring(4) + ", The";
      else
        return name;
    }

    public int Compare(Cave x, Cave y)
    {
      return SortName(x).CompareTo(SortName(y));
    }
  }
}
