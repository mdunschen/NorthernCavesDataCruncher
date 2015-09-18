using System;
using System.IO;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Diagnostics;
using System.Xml.Schema;

namespace NCDataCruncher
{
  public partial class Form1 : Form
  {
    private const bool SplitMinorCaves = false;

    public class Prefs
    {
      public string Folder { get; set; }
      public string Username { get; set; }
      public string Password { get; set; }
    }

    private static readonly string ExeFolder = Path.GetDirectoryName(Application.ExecutablePath);
    private static readonly string AppDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
    private static string ExeName = Path.GetFileNameWithoutExtension(Application.ExecutablePath);
    private static readonly string SettingsFolder = Path.Combine(AppDataFolder, Path.Combine("Northern Caves", ExeName));
    private static readonly string SettingsFileName = "Settings.xml";
    private static readonly string SettingsFilePath = Path.Combine(SettingsFolder, SettingsFileName);

    private const string BaseFont = "Calibri";
    private const string AccentColour = "485A91";

    private const int BaseFontSize = 20; // Points x 2
    private const int SmallBaseFontSize = 18;
    private const int AreaHeaderFontSize = 48;
    private const int CaveNameFontSize = 34;
    private const int MinorCaveNameFontSize = 24;
    private const int GradeFontSize = 28;
    private const int SubHeaderFontSize = 24;
    private const int SubHeader2FontSize = 24;
    private const int CaveDataFontSize = 22;

    private const double PageWidthCm = 12;
    private const double CaveDataTab1 = 3.5;
    private const double CaveDataTab2 = 7.5;

    private const string urlString = @"http://nc.ulsa.org.uk/";
    private const int Vol1 = 0;
    private const int Vol2 = 1;
    private const int Vol3 = 2;
    //private const int MLC = 3;
    private const int Test = 3;
    private const int bookCount = 4;

    private Prefs prefs;
    private Dictionary<string, int>[] areaNumbers;
    private Dictionary<string, List<string>> MajorCaves;
    private string authToken;
    private bool loggedIn;
    private bool updating;

    public Form1()
    {
      InitializeComponent();
      cmbBook.Items.Add("Northern Caves 1 - Wharfdale & the North-East");
      cmbBook.Items.Add("Northern Caves 2 - The Three Peaks");
      cmbBook.Items.Add("Northern Caves 3 - The Three Counties System and the North West");
      cmbBook.Items.Add("Test");
      var serializer = new XmlSerializer(typeof(Prefs));
      if (File.Exists(SettingsFilePath))
        using (Stream stream = File.OpenRead(SettingsFilePath))
          prefs = (Prefs)serializer.Deserialize(stream);
      else
        prefs = new Prefs();
      txtFolder.Text = prefs.Folder;

      areaNumbers = new Dictionary<string, int>[bookCount];
      for (int i = 0; i < bookCount; i++)
        areaNumbers[i] = new Dictionary<string, int>();

      // Volume 3
      areaNumbers[Vol3].Add("Scales Moor", 40);
      areaNumbers[Vol3].Add("East Kingsdale", 15);
      areaNumbers[Vol3].Add("Kingsdale Head", 22);
      areaNumbers[Vol3].Add("West Kingsdale", 51);
      areaNumbers[Vol3].Add("Marble Steps", 28);
      areaNumbers[Vol3].Add("Leck Fell", 24);
      areaNumbers[Vol3].Add("Ease Gill", 14);
      areaNumbers[Vol3].Add("Barbondale", 4);
      areaNumbers[Vol3].Add("Dentdale", 13);
      areaNumbers[Vol3].Add("Garsdale and Great Knoutberry", 59);
      areaNumbers[Vol3].Add("Wild Boar Fell", 53);
      areaNumbers[Vol3].Add("Mallerstang", 27);
      areaNumbers[Vol3].Add("Brough", 7);
      areaNumbers[Vol3].Add("Vale of Eden", 48);
      areaNumbers[Vol3].Add("Bowland", 6);
      areaNumbers[Vol3].Add("Morecambe Bay", 29);
      areaNumbers[Vol3].Add("Other Areas", 35);

      // Volume 2
      areaNumbers[Vol2].Add("Malham", 26);
      areaNumbers[Vol2].Add("Stockdale", 41);
      areaNumbers[Vol2].Add("Attermire", 3);
      areaNumbers[Vol2].Add("Giggleswick Scar", 18);
      areaNumbers[Vol2].Add("Fountains Fell", 16);
      areaNumbers[Vol2].Add("Penyghent", 37);
      areaNumbers[Vol2].Add("Birkwith", 5);
      areaNumbers[Vol2].Add("Ribblehead", 39);
      areaNumbers[Vol2].Add("Alum Pot", 2);
      areaNumbers[Vol2].Add("Moughton", 30);
      areaNumbers[Vol2].Add("The Allotment", 45);
      areaNumbers[Vol2].Add("Gaping Gill", 17);
      areaNumbers[Vol2].Add("Newby Moss", 31);
      areaNumbers[Vol2].Add("White Scar", 52);
      areaNumbers[Vol2].Add("Chapel-Le-Dale", 9);
      areaNumbers[Vol2].Add("Park Fell", 36);
      areaNumbers[Vol2].Add("Bruntscar", 8);

      // Volume 1
      areaNumbers[Vol1].Add("Grassington", 19);
      areaNumbers[Vol1].Add("Great Whernside", 20);
      areaNumbers[Vol1].Add("Upper Wharfedale", 47);
      areaNumbers[Vol1].Add("Langstrothdale", 23);
      areaNumbers[Vol1].Add("Lower Littondale", 25);
      areaNumbers[Vol1].Add("Darnbrook and Cowside", 12);
      areaNumbers[Vol1].Add("Upper Littondale", 46);
      areaNumbers[Vol1].Add("Penyghent Gill", 38);
      areaNumbers[Vol1].Add("Cosh and Foxup", 10);
      areaNumbers[Vol1].Add("Stump Cross", 42);
      areaNumbers[Vol1].Add("Nidderdale", 32);
      areaNumbers[Vol1].Add("Coverdale and Bishopdale", 11);
      areaNumbers[Vol1].Add("Wensleydale", 50);
      areaNumbers[Vol1].Add("Swaledale and Arkengarthdale", 43);
      areaNumbers[Vol1].Add("Gretadale",21 );
      areaNumbers[Vol1].Add("Teesdale", 44);
      areaNumbers[Vol1].Add("Weardale", 49);
      areaNumbers[Vol1].Add("Alston", 1);
      areaNumbers[Vol1].Add("Northumberland", 34);
      areaNumbers[Vol1].Add("North York Moors", 33);
      areaNumbers[Vol1].Add("Other Caves", 55);

      // Moorland Caver
      //areaNumbers[MLC].Add("MLC - County Durham", 54);
      //areaNumbers[MLC].Add("MLC - Derbyshire and Nottinghamshire", 58);
      //areaNumbers[MLC].Add("MLC - North and West Yorkshire", 56);
      //areaNumbers[MLC].Add("MLC - South Yorkshire", 57);

      // Test
      areaNumbers[Test].Add("Test", 999);

    }

    private void btnBrowse_Click(object sender, EventArgs e)
    {
      var dlg = new FolderBrowserDialog();
      if (dlg.ShowDialog() == DialogResult.OK)
        txtFolder.Text = dlg.SelectedPath;
    }

    private void tree_AfterCheck(object sender, TreeViewEventArgs e)
    {
      if (updating)
        return;
      updating = true;
      if (e.Node.Text == "All")
        for (int i = 1; i < tree.Nodes.Count; i++)
          tree.Nodes[i].Checked = tree.Nodes[0].Checked;
      else if (!e.Node.Checked)
        tree.Nodes[0].Checked = false;
      updating = false;
    }

    private void btnLogin_Click(object sender, EventArgs e)
    {
      var dlg = new DlgLogin(urlString);
      dlg.lblLoginTo.Text = "Log in to " + urlString;
      dlg.txtUsername.Text = prefs.Username;
      dlg.txtPassword.Text = prefs.Password;
      loggedIn = dlg.ShowDialog() == DialogResult.OK;
      //lblLoggedIn.Text = loggedIn ? "Logged in" : "Not logged in";
      if (loggedIn)
      {
        authToken = dlg.AuthToken;
        prefs.Username = dlg.txtUsername.Text;
        prefs.Password = dlg.txtPassword.Text;
      }
      btnLogin.Enabled = false;
      btnDownloadXml.Enabled = true;
      btnScreenScrape.Enabled = true;
    }

    private void btnDownloadXml_Click(object sender, EventArgs e)
    {
      string folder = Path.Combine(txtFolder.Text, "xml");
      if (!Directory.Exists(folder))
        Directory.CreateDirectory(folder);
      foreach (TreeNode node in tree.Nodes)
        if (node.Text != "All" && node.Checked)
          GetXml(node.Text, folder);
      scroller.AddStr("");
    }

    private void btnScreenScrape_Click(object sender, EventArgs e)
    {
      string folder = Path.Combine(txtFolder.Text, "xml");
      if (!Directory.Exists(folder))
      {
        scroller.AddStr("Creating folder: " + folder);
        Directory.CreateDirectory(folder);
      }
      foreach (TreeNode node in tree.Nodes)
        if (node.Text != "All" && node.Checked)
        {
          scroller.AddStr("Screen Scraping " + node.Text);
          string filePath = Path.Combine(folder, node.Text + ".xml");
          ScreenScrape(filePath);
        }
      scroller.AddStr("");
    }

    private Dictionary<string, List<string>> LoadMajorCaves(string filePath)
    {
      var majorCaves = new Dictionary<string, List<string>>();
      var xDoc = XDocument.Load(filePath);
      var root = xDoc.Element("majorcaves");
      foreach (var areaElement in root.Elements())
      {
        string area = areaElement.Attribute("name").Value;
        majorCaves.Add(area, new List<string>());
        foreach (var caveElement in areaElement.Elements())
          majorCaves[area].Add(caveElement.Value);
      }
      return majorCaves;
    }

    private void btnCreateDocx_Click(object sender, EventArgs e)
    {
      if (SplitMinorCaves)
        MajorCaves = LoadMajorCaves(Path.Combine(Path.Combine(txtFolder.Text, "xml"), "Major Caves.xml"));

      string outputFolder = Path.Combine(txtFolder.Text, "doc");
      string inputFolder = Path.Combine(txtFolder.Text, "xml");
      string sampleFolder = Path.Combine(txtFolder.Text, "template");
      if (!Directory.Exists(outputFolder))
      {
        scroller.AddStr("Creating folder: " + outputFolder);
        Directory.CreateDirectory(outputFolder);
      }
      scroller.AddStr("Creating DocX Files");
      foreach (TreeNode node in tree.Nodes)
        if (node.Text != "All" && node.Checked)
        {
          string outputFilePath = Path.Combine(outputFolder, node.Text + ".docx");
          scroller.AddStr("  " + node.Text);
          Func<string, string, bool> majorCave;
          if (SplitMinorCaves)
            majorCave  = (area, cave) => MajorCaves[area].Contains(cave);
          else
            majorCave = (area, cave) => true;
          //CreateDocxFile(Path.Combine(inputFolder, node.Text + ".xml"), Path.Combine(outputFolder, node.Text + ".docx"), majorCave );
          CreateDocx(Path.Combine(inputFolder, node.Text + ".xml"), outputFilePath, Path.Combine(sampleFolder, "Area Template.docx"));
          if (chkOpenWhenDone.Checked)
            Process.Start(outputFilePath);
        }
      scroller.AddStr("");
    }

    private void GetXml(string area, string folder)
    {
      scroller.AddStr(area);
      int areaId = areaNumbers[cmbBook.SelectedIndex][area];
      string fullUrlString = string.Format("{0}create_xml.php?area_id={1}", urlString, areaId);
      HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(new Uri(fullUrlString));
      var cookies = new CookieContainer();
      var cookie = new Cookie("nc_cookie", authToken);
      cookies.Add(new Uri(urlString), cookie);
      webRequest.CookieContainer = cookies;
      var response = webRequest.GetResponse();
      Stream stream = response.GetResponseStream();
      //var reader = new StreamReader(stream, Encoding.UTF8);
      //string content = reader.ReadToEnd();
      var xDoc = XDocument.Load(stream);
      xDoc.Save(Path.Combine(folder,area + ".xml"));
    }

    private string GetCaveDescription(string id)
    {
      string fullUrlString = string.Format("{0}edit_text.php?cave_id={1}", urlString, id);
      HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(new Uri(fullUrlString));
      var cookies = new CookieContainer();
      var cookie = new Cookie("nc_cookie", authToken);
      cookies.Add(new Uri(urlString), cookie);
      webRequest.CookieContainer = cookies;
      var response = webRequest.GetResponse();
      Stream stream = response.GetResponseStream();
      var reader = new StreamReader(stream, Encoding.UTF8);
      string content = reader.ReadToEnd();

      string startTag = "id=\"textarea\">";
      string endTag = "</textarea>";
      int startIndex = content.IndexOf(startTag);
      int endIndex = content.IndexOf(endTag);
      if (startIndex > 0)
        return content.Substring(startIndex + startTag.Length, endIndex - startIndex - startTag.Length);
      return null;
    }

    private void ScreenScrape(string filePath)
    {
      var xDoc = XDocument.Load(filePath);
      foreach (var caveElement in xDoc.Element("area").Elements("cave"))
      {
        string id = caveElement.Element("id").Value;
        string name = caveElement.Element("name").Value;
        scroller.AddStr("   " + id + "  " + name);
        string description = GetCaveDescription(id);
        var descElement = XElement.Parse("<description>" + description + "</description>");
        var descriptionElement = caveElement.Element("description");
        if (descriptionElement != null)
          descriptionElement.Remove();
        caveElement.Add(descElement);
        int? upperGrade;
        string sortName;
        GetMissingFields(id, out upperGrade, out sortName);
        if (upperGrade.HasValue)
        {
          XElement element = caveElement.Element("uppergrade");
          if (element == null)
          {
            element = new XElement("uppergrade", upperGrade);
            caveElement.Add(element);
          }
          else
            element.Value = upperGrade.ToString();
        }
        if (sortName != null && sortName != "")
        {
          XElement element = caveElement.Element("Sortname");
          if (element == null)
          {
            element = new XElement("sortname", sortName);
            caveElement.Add(element);
          }
          else
            element.Value = sortName;
        }
      }
      xDoc.Save(filePath);
      scroller.AddStr("");
    }

    private void GetMissingFields(string id, out int? upperGrade, out string sortName)
    {
      string fullUrlString = string.Format("{0}edit_record.php?cave_id={1}", urlString, id);
      HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(new Uri(fullUrlString));
      var cookies = new CookieContainer();
      var cookie = new Cookie("nc_cookie", authToken);
      cookies.Add(new Uri(urlString), cookie);
      webRequest.CookieContainer = cookies;
      var response = webRequest.GetResponse();
      Stream stream = response.GetResponseStream();
      var reader = new StreamReader(stream, Encoding.UTF8);
      string content = reader.ReadToEnd();

      int baseIndex = content.IndexOf("<td class=\"k\">Upper grade</td>");
      int index1 = content.IndexOf("value=\"", baseIndex) + 7;
      int index2 = content.IndexOf("\"", index1);
      string upperGradeString = content.Substring(index1, index2 - index1);
      if (upperGradeString != "")
        upperGrade = int.Parse(upperGradeString);
      else
        upperGrade = null;

      baseIndex = content.IndexOf("Sort name</td>");
      index1 = content.IndexOf("value=\"", baseIndex) + 7;
      index2 = content.IndexOf("\"", index1);
      sortName = content.Substring(index1, index2 - index1);      
    }

    private void Form1_Load(object sender, EventArgs e)
    {
      cmbBook.SelectedIndex = 2;
    }

    private void cmbBook_SelectedIndexChanged(object sender, EventArgs e)
    {
      tree.Nodes.Clear();
      tree.Nodes.Add("All");
      foreach (string area in areaNumbers[cmbBook.SelectedIndex].Keys)
        tree.Nodes.Add(area);
    }

    private void btnGetCaveList_Click(object sender, EventArgs e)
    {
      scroller.AddStr("Create Cave List");
      var outputXDoc = new XDocument();
      var root = new XElement("cavelist");
      outputXDoc.Add(root);
      foreach (TreeNode node in tree.Nodes)
        if (node.Text != "All" && node.Checked)
        {
          scroller.AddStr("  " + node.Text);
          GetAreaCaveList(node.Text, outputXDoc, root);
          scroller.AddStr("");
        }
      outputXDoc.Save(Path.Combine(txtFolder.Text, Path.Combine("xml","CaveList.xml")));
    }

    private void GetAreaCaveList(string area, XDocument outputXDoc, XElement outputRoot)
    {
        string filePath = Path.Combine(txtFolder.Text, Path.Combine("xml", area + ".xml"));
        var xDoc = XDocument.Load(filePath);
        XElement areaElem = new XElement("area");
        outputRoot.Add(areaElem);
        areaElem.Add(new XAttribute("Area", area));
        //areaElem.Add(new XAttribute("Sequence", i.ToString()));
        foreach (var caveElement in xDoc.Element("area").Elements("cave"))
        {
          string id = caveElement.Element("id").Value;
          string name = caveElement.Element("name").Value;
          XElement aliasElem = caveElement.Element("alias");
          XElement gridElem = caveElement.Element("grid");
          XElement latElem = caveElement.Element("latitude");
          XElement lonElem = caveElement.Element("longitude");
          XElement elevationElem = caveElement.Element("elevation");
          XElement lengthElem = caveElement.Element("length");
          XElement lengthNoteElem = caveElement.Element("length_note");
          XElement depthElem = caveElement.Element("depth");
          XElement depthNoteElem = caveElement.Element("depth_note");
          XElement gradeElem = caveElement.Element("grade");
          XElement upperGradeElem = caveElement.Element("uppergrade");
          XElement warningElem = caveElement.Element("warning");
          XElement permissionElem = caveElement.Element("permission");
          XElement rockElem = caveElement.Element("rock");
          XElement explorationElem = caveElement.Element("exploration");
          XElement hydrologyElem = caveElement.Element("hydrology");
          XElement sortnameElem = caveElement.Element("sortname");

          string alias = aliasElem!=null?aliasElem.Value:null;
          string grid = gridElem!=null?gridElem.Value:null;
          string latString = latElem!=null?latElem.Value:null;
          string lonString = lonElem!=null?lonElem.Value:null;
          string elevationString = elevationElem!=null?elevationElem.Value:null;
          string lengthString = lengthElem!=null?lengthElem.Value:null;
          string lengthNote = lengthNoteElem!=null?lengthNoteElem.Value:null;
          string depthString = depthElem!=null?depthElem.Value:null;
          string depthNote = depthNoteElem!=null?depthNoteElem.Value:null;
          string gradeString = gradeElem!=null?gradeElem.Value:null;
          string upperGradeString = upperGradeElem!=null?upperGradeElem.Value:null;
          string warning = warningElem!=null?warningElem.Value:null;
          string permission = permissionElem!=null?permissionElem.Value:null;
          string rock = rockElem!=null?rockElem.Value:null;
          string exploration = explorationElem!=null?explorationElem.Value:null;
          string hydrology = hydrologyElem != null ? hydrologyElem.Value : null;
          string sortname = sortnameElem != null ? sortnameElem.Value : null;

          double? lat = null;
          if (latString != null)
            lat = double.Parse(latString);

          double? lon = null;
          if (lonString != null)
            lon = double.Parse(lonString);

          double? elevation = null;
          if (elevationString != null)
            elevation = double.Parse(elevationString);

          double? length = null;
          if (lengthString != null)
            length = double.Parse(lengthString);

          double? depth = null;
          if (depthString != null)
            depth = double.Parse(depthString);

          int? grade = null;
          if (gradeString != null)
            grade = int.Parse(gradeString);

          int? upperGrade = null;
          if (upperGradeString != null)
            upperGrade = int.Parse(upperGradeString);

          scroller.AddStr("    " + name);
          var outputElem = new XElement("cave", new XAttribute("Name", name), new XAttribute("Id", id));

          if (alias != null) outputElem.Add(new XAttribute("Alias", alias));
          if (grid != null) outputElem.Add(new XAttribute("Grid", grid));
          if (lat.HasValue) outputElem.Add(new XAttribute("Lat", lat.Value));
          if (lon.HasValue) outputElem.Add(new XAttribute("Lon", lon.Value));
          if (elevation.HasValue) outputElem.Add(new XAttribute("Elevation", elevation.Value));
          if (length.HasValue) outputElem.Add(new XAttribute("Length", length.Value));
          if (lengthNote != null) outputElem.Add(new XAttribute("LengthNote", lengthNote));
          if (depth.HasValue) outputElem.Add(new XAttribute("Depth", depth.Value));
          if (depthNote != null) outputElem.Add(new XAttribute("DepthNote", depthNote));
          if (grade.HasValue) outputElem.Add(new XAttribute("Grade", grade.Value));
          if (upperGrade.HasValue) outputElem.Add(new XAttribute("UpperGrade", upperGrade.Value));
          if (warning != null) outputElem.Add(new XAttribute("Warning", warning));
          if (permission != null) outputElem.Add(new XAttribute("Permission", permission));          
          if (rock != null) outputElem.Add(new XAttribute("Rock", rock));
          if (exploration != null) outputElem.Add(new XAttribute("Exploration", exploration));
          if (hydrology != null) outputElem.Add(new XAttribute("Hydrology", hydrology));
          if (sortname != null) outputElem.Add(new XAttribute("SortName", sortname));
          
          if (length >= 100 || depth >= 30)
            outputElem.Add(new XAttribute("Major", "Y"));
          areaElem.Add(outputElem);
        }
      }

    private void Form1_FormClosing(object sender, FormClosingEventArgs e)
    {
      var serializer = new XmlSerializer(typeof(Prefs));
      if (!Directory.Exists(SettingsFolder))
        Directory.CreateDirectory(SettingsFolder);
      using (Stream stream = File.Create(SettingsFilePath))
        serializer.Serialize(stream, prefs);
    }

    private void txtFolder_TextChanged(object sender, EventArgs e)
    {
      prefs.Folder = txtFolder.Text;
    }

    /// <summary>
    /// Convert cm to subpoints (a subpoint is 20 dtp points or 1/1440 inch
    /// </summary>
    private int CmToSubPoints(double cm)
    {
      const double cmToInch = 2.54;
      const double PointsPerInch = 72;
      return (int)Math.Round(cm * PointsPerInch * 20 / cmToInch);
    }

    private void CreateDocxFile(string inputFilePath, string outputFilePath, Func<string, string, bool> majorCave)
    {

      //var area = Area.Load(inputFilePath);
      //using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputFilePath, WordprocessingDocumentType.Document))
      //{
      //  MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
      //  StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();

      //  var locationAndLengthTabs = new Tabs(
      //    new TabStop { Val = TabStopValues.Left, Position = CmToSubPoints(CaveDataTab1) },
      //    new TabStop { Val = TabStopValues.Left, Position = CmToSubPoints(CaveDataTab2) },
      //    new TabStop { Val = TabStopValues.Right, Position = CmToSubPoints(PageWidthCm) });

      //  var minorLocationAndLengthTabs = new Tabs(
      //    new TabStop { Val = TabStopValues.Left, Position = CmToSubPoints(CaveDataTab1) },
      //    new TabStop { Val = TabStopValues.Left, Position = CmToSubPoints(CaveDataTab2) },
      //    new TabStop { Val = TabStopValues.Right, Position = CmToSubPoints(PageWidthCm) });

      //  var gradeTabs = new Tabs(new TabStop { Val = TabStopValues.Right, Position = CmToSubPoints(PageWidthCm) });
      //  var minorGradeTabs = new Tabs(new TabStop { Val = TabStopValues.Right, Position = CmToSubPoints(PageWidthCm) });
      //  var subHeaderGradeTabs = new Tabs(new TabStop { Val = TabStopValues.Right, Position = CmToSubPoints(PageWidthCm) });
      //  var subHeader2GradeTabs = new Tabs(new TabStop { Val = TabStopValues.Right, Position = CmToSubPoints(PageWidthCm) });

      //  stylesPart.Styles = new Styles();

      //  //Paragraph Styles
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.AreaHeader, StyleValues.Paragraph, BaseFont, AreaHeaderFontSize, hexColour: AccentColour, bold: true));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.CaveName, StyleValues.Paragraph, BaseFont, CaveNameFontSize, hexColour: AccentColour, tabs: gradeTabs));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.Grade, StyleValues.Character, null, GradeFontSize));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.SubHeader, StyleValues.Paragraph, BaseFont, SubHeaderFontSize, bold: true, hexColour: AccentColour, tabs: subHeaderGradeTabs));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.SubHeader2, StyleValues.Paragraph, BaseFont, SubHeader2FontSize, bold: true, tabs: subHeader2GradeTabs));

      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.Alias, StyleValues.Paragraph, BaseFont, CaveDataFontSize, bold: true, italic: true));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.LocationAndLength, StyleValues.Paragraph, BaseFont, CaveDataFontSize, bold: true, italic: true, tabs: locationAndLengthTabs));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.ExplorationHistory, StyleValues.Paragraph, BaseFont, BaseFontSize, italic: true));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.Permission, StyleValues.Paragraph, BaseFont, BaseFontSize, italic: true));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.Warning, StyleValues.Paragraph, BaseFont, BaseFontSize, bold: true));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.LengthDepthNote, StyleValues.Paragraph, BaseFont, BaseFontSize));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.Rock, StyleValues.Paragraph, BaseFont, BaseFontSize, bold: true));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.Hydrology, StyleValues.Paragraph, BaseFont, BaseFontSize, bold: true));

      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.CaveDescription, StyleValues.Paragraph, BaseFont, BaseFontSize));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.OffRoute, StyleValues.Character, italic: true));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.Sump, StyleValues.Character, bold: true));

      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.PitchTableHeader, StyleValues.Paragraph, BaseFont, BaseFontSize, bold: true));

      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.Tackle, StyleValues.Character, BaseFont, BaseFontSize, bold: true));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.TackleHeader, StyleValues.Paragraph, BaseFont, BaseFontSize, bold: true));
      //  stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.TackleTable, StyleValues.Table, BaseFont));

      //  if (SplitMinorCaves)
      //  {
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorCavesHeader, StyleValues.Paragraph, BaseFont, CaveNameFontSize, hexColour: AccentColour));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorCaveName, StyleValues.Paragraph, BaseFont, MinorCaveNameFontSize, hexColour: AccentColour, tabs: minorGradeTabs));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorAlias, StyleValues.Paragraph, BaseFont, SmallBaseFontSize, bold: true, italic: true));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorLocationAndLength, StyleValues.Paragraph, BaseFont, SmallBaseFontSize, bold: true, italic: true, tabs: minorLocationAndLengthTabs));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorExplorationHistory, StyleValues.Paragraph, BaseFont, SmallBaseFontSize, italic: true));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorPermission, StyleValues.Paragraph, BaseFont, SmallBaseFontSize, italic: true));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorWarning, StyleValues.Paragraph, BaseFont, SmallBaseFontSize, bold: true));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorLengthDepthNote, StyleValues.Paragraph, BaseFont, SmallBaseFontSize));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorRock, StyleValues.Paragraph, BaseFont, SmallBaseFontSize, bold: true));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorHydrology, StyleValues.Paragraph, BaseFont, SmallBaseFontSize, bold: true));

      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorCaveDescription, StyleValues.Paragraph, BaseFont, SmallBaseFontSize));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorOffRoute, StyleValues.Character, italic: true));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorSump, StyleValues.Character, bold: true));

      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorTackle, StyleValues.Character, BaseFont, SmallBaseFontSize, bold: true));
      //    stylesPart.Styles.Append(Docx.CreateStyle(StyleNames.MinorTackleHeader, StyleValues.Paragraph, BaseFont, SmallBaseFontSize, bold: true));
      //  }


      //  stylesPart.Styles.Save();

      //  mainPart.Document = new Document();
      //  Body body = new Body();

      //  var sectionProps = new SectionProperties();
      //  var pageSize = new PageSize() { Width = (UInt32Value)8391U, Height = (UInt32Value)11907U, Code = (UInt16Value)11U };
      //  var pageMargin = new PageMargin() { Top = 720, Right = (UInt32Value)720U, Bottom = 720, Left = (UInt32Value)720U, Header = (UInt32Value)708U, Footer = (UInt32Value)708U, Gutter = (UInt32Value)0U };
      //  //var columns = new Columns() { Space = "708" };
      //  //var docGrid = new DocGrid() { LinePitch = 360 };

      //  sectionProps.Append(pageSize);
      //  sectionProps.Append(pageMargin);
      //  //sectionProps.Append(columns);
      //  //sectionProps.Append(docGrid);
      //  body.Append(sectionProps);

      //  area.WriteDocx(body, majorCave);
      //  mainPart.Document.Append(body);
      //  mainPart.Document.Save();
      //}
    }

    private void CreateDocx(string inputFilePath, string outputFilePath, string sampleFilePath)
    {
      try
      {
        var area = Area.Load(inputFilePath);
        File.Copy(sampleFilePath, outputFilePath, true);
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputFilePath, true))
        {
          Body body = wordDoc.MainDocumentPart.Document.Body;
          SectionProperties sectionProps = null;
          foreach (var element in body.Elements())
            if (element is SectionProperties)
              sectionProps = element as SectionProperties;
          body.RemoveAllChildren();
          body.Append(sectionProps);
          area.WriteDocx(body, (a, b) => true);
          wordDoc.MainDocumentPart.Document.Save();
        }
      }
      catch (Exception ex)
        {
          MessageBox.Show(ex.Message);
        }
    }

    private void btnValidateXml_Click(object sender, EventArgs e)
    {
      string xmlFolder = Path.Combine(txtFolder.Text, "xml");
      string xsdFilePath = Path.Combine(Path.Combine(txtFolder.Text, "template"), "nc.xsd");
      foreach (TreeNode node in tree.Nodes)
        if (node.Text != "All" && node.Checked)
        {
          scroller.AddStr("Validating " + node.Text + ".xml");
          string filePath = Path.Combine(xmlFolder, node.Text + ".xml");
          string errorMessage;
          if (ValidateXmlWithXsd(filePath, xsdFilePath, out errorMessage))
            scroller.AddStr("  OK");
          else
            scroller.AddStr("  Failed: " + errorMessage);
        }
      scroller.AddStr("");
   
    }

    private static bool ValidateXmlWithXsd(string xmlUri, string xsdUri, out string errorMessage)
    {
      errorMessage = null;
      try
      {
        XmlReaderSettings xmlSettings = new XmlReaderSettings();
        xmlSettings.Schemas = new System.Xml.Schema.XmlSchemaSet();
        xmlSettings.Schemas.Add(null, xsdUri);
        xmlSettings.ValidationType = ValidationType.Schema;
        XmlReader reader = XmlReader.Create(xmlUri, xmlSettings);
        while (reader.Read())
          ;
        return true;
      }
      catch (XmlSchemaValidationException ex)
      {
        errorMessage = ex.Message;
        return false;
      }
    }


    //private void button1_Click(object sender, EventArgs e)
    //{
    //  //string outputFolder = Path.Combine(txtFolder.Text, "doc");
    //  //string filePath = Path.Combine(outputFolder, "Kingsdale Head" + ".docx");
    //  string folder = @"C:\Users\Admin\Documents\Northern Caves 2\SVN Working Copy\Volume 3\Template";
    //  string filePath = Path.Combine(folder, "sample.docx");
    //  string outputFilePath = Path.Combine(folder, "output.docx");
    //  File.Copy(filePath, outputFilePath, true);
    //  using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputFilePath, true))
    //  {
    //    Body body = wordDoc.MainDocumentPart.Document.Body;
    //    SectionProperties sectionProps = null;
    //    foreach (var element in body.Elements())
    //      if (element is SectionProperties)
    //        sectionProps = element as SectionProperties;
    //    body.RemoveAllChildren();
    //    body.Append(sectionProps);
    //    body.Append(new Paragraph(new Run(new Text("Hello World"))));
    //    wordDoc.MainDocumentPart.Document.Save();

    //  }
    //}
  }
}
