using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NCDataCruncher
{
  public class Scroller : ListBox
  {
    private int maxLines = 1000;

    private delegate void MessageDelegate(String str);

    public int MaxLines
    {
      get { return this.maxLines; }
      set { maxLines = value; purge(); }
    }

    public void AddStr(string str)
    {
      if (InvokeRequired)
        BeginInvoke(new MessageDelegate(AddString), new object[] { str });
      else
        AddString(str);
    }

    private void AddString(String str)
    {
      if (str == null)
        return;
      while (Items.Count > maxLines)
        Items.RemoveAt(0);
      Items.Add(str);
      SelectedIndex = Items.Count - 1;
    }

    private void purge()
    {
      while (Items.Count > maxLines)
        Items.RemoveAt(0);
      SelectedIndex = Items.Count - 1;
    }
  }
}
