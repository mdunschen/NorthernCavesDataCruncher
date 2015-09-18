using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;

namespace NCDataCruncher
{
  public partial class DlgLogin : Form
  {
    private string urlString;

    public string AuthToken { get; private set; }
    // bool LoggedIn { get; private set;}

    public DlgLogin()
    {
      InitializeComponent();
    }

    public DlgLogin(string urlString)
      : this()
    {
      this.urlString = urlString;
    }

    private void btnOK_Click(object sender, EventArgs e)
    {
      bool LoggedIn;
      lblFailed.Visible = false;
      try
      {
        AuthToken = Login(urlString, txtUsername.Text, txtPassword.Text);
        LoggedIn = true;
        DialogResult = DialogResult.OK;
      }
      catch (Exception ex)
      {
        LoggedIn = false;
        lblFailed.Visible = true;
      }
    }

    private string Login(string urlString, string username, string password)
    {
      var uri = new Uri(urlString);
      var request = (HttpWebRequest)WebRequest.Create(uri);
      request.CookieContainer = new CookieContainer();
      request.Method = "POST";
      request.ContentType = "application/x-www-form-urlencoded";
      using (var requestStream = request.GetRequestStream())
      using (var writer = new StreamWriter(requestStream))
        writer.Write(string.Format("A_USER={0}&passwd={1}", username, password));
      HttpWebResponse response = (HttpWebResponse)request.GetResponse();
      string authToken = response.Cookies["nc_cookie"].Value;
      return authToken;
    }
  }
}
