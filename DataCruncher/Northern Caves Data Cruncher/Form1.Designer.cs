namespace NCDataCruncher
{
  partial class Form1
  {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
      if (disposing && (components != null))
      {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("All");
            this.btnDownloadXml = new System.Windows.Forms.Button();
            this.btnLogin = new System.Windows.Forms.Button();
            this.tree = new System.Windows.Forms.TreeView();
            this.btnScreenScrape = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFolder = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnCreateDocx = new System.Windows.Forms.Button();
            this.cmbBook = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnGetCaveList = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.chkOpenWhenDone = new System.Windows.Forms.CheckBox();
            this.btnValidateXml = new System.Windows.Forms.Button();
            this.scroller = new NCDataCruncher.Scroller();
            this.SuspendLayout();
            // 
            // btnDownloadXml
            // 
            this.btnDownloadXml.Enabled = false;
            this.btnDownloadXml.Location = new System.Drawing.Point(12, 152);
            this.btnDownloadXml.Name = "btnDownloadXml";
            this.btnDownloadXml.Size = new System.Drawing.Size(183, 23);
            this.btnDownloadXml.TabIndex = 3;
            this.btnDownloadXml.Text = "2 Download XML";
            this.btnDownloadXml.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnDownloadXml.UseVisualStyleBackColor = true;
            this.btnDownloadXml.Click += new System.EventHandler(this.btnDownloadXml_Click);
            // 
            // btnLogin
            // 
            this.btnLogin.Location = new System.Drawing.Point(12, 123);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(183, 23);
            this.btnLogin.TabIndex = 2;
            this.btnLogin.Text = "1 Log In";
            this.btnLogin.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnLogin.UseVisualStyleBackColor = true;
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // tree
            // 
            this.tree.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.tree.CheckBoxes = true;
            this.tree.Location = new System.Drawing.Point(212, 123);
            this.tree.Name = "tree";
            treeNode1.Name = "Node0";
            treeNode1.Text = "All";
            this.tree.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1});
            this.tree.ShowLines = false;
            this.tree.ShowNodeToolTips = true;
            this.tree.ShowPlusMinus = false;
            this.tree.ShowRootLines = false;
            this.tree.Size = new System.Drawing.Size(244, 369);
            this.tree.TabIndex = 6;
            this.tree.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.tree_AfterCheck);
            // 
            // btnScreenScrape
            // 
            this.btnScreenScrape.Enabled = false;
            this.btnScreenScrape.Location = new System.Drawing.Point(12, 181);
            this.btnScreenScrape.Name = "btnScreenScrape";
            this.btnScreenScrape.Size = new System.Drawing.Size(183, 23);
            this.btnScreenScrape.TabIndex = 4;
            this.btnScreenScrape.Text = "3 ScreenScrape Cave Descriptions";
            this.btnScreenScrape.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnScreenScrape.UseVisualStyleBackColor = true;
            this.btnScreenScrape.Click += new System.EventHandler(this.btnScreenScrape_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 74);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Root Folder";
            // 
            // txtFolder
            // 
            this.txtFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtFolder.Location = new System.Drawing.Point(12, 90);
            this.txtFolder.Name = "txtFolder";
            this.txtFolder.Size = new System.Drawing.Size(363, 20);
            this.txtFolder.TabIndex = 1;
            this.txtFolder.TextChanged += new System.EventHandler(this.txtFolder_TextChanged);
            // 
            // btnBrowse
            // 
            this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowse.Location = new System.Drawing.Point(381, 88);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 0;
            this.btnBrowse.Text = "Browse ...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnCreateDocx
            // 
            this.btnCreateDocx.Location = new System.Drawing.Point(12, 210);
            this.btnCreateDocx.Name = "btnCreateDocx";
            this.btnCreateDocx.Size = new System.Drawing.Size(183, 23);
            this.btnCreateDocx.TabIndex = 5;
            this.btnCreateDocx.Text = "4 Create DocX files";
            this.btnCreateDocx.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCreateDocx.UseVisualStyleBackColor = true;
            this.btnCreateDocx.Click += new System.EventHandler(this.btnCreateDocx_Click);
            // 
            // cmbBook
            // 
            this.cmbBook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBook.FormattingEnabled = true;
            this.cmbBook.Location = new System.Drawing.Point(12, 38);
            this.cmbBook.Name = "cmbBook";
            this.cmbBook.Size = new System.Drawing.Size(363, 21);
            this.cmbBook.TabIndex = 9;
            this.cmbBook.SelectedIndexChanged += new System.EventHandler(this.cmbBook_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(32, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Book";
            // 
            // btnGetCaveList
            // 
            this.btnGetCaveList.Location = new System.Drawing.Point(12, 349);
            this.btnGetCaveList.Name = "btnGetCaveList";
            this.btnGetCaveList.Size = new System.Drawing.Size(183, 23);
            this.btnGetCaveList.TabIndex = 11;
            this.btnGetCaveList.Text = "Create Cave List";
            this.btnGetCaveList.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnGetCaveList.UseVisualStyleBackColor = true;
            this.btnGetCaveList.Click += new System.EventHandler(this.btnGetCaveList_Click);
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.DarkBlue;
            this.label3.Location = new System.Drawing.Point(9, 478);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 13);
            this.label3.TabIndex = 12;
            this.label3.Text = "V1.7";
            // 
            // chkOpenWhenDone
            // 
            this.chkOpenWhenDone.AutoSize = true;
            this.chkOpenWhenDone.Location = new System.Drawing.Point(12, 239);
            this.chkOpenWhenDone.Name = "chkOpenWhenDone";
            this.chkOpenWhenDone.Size = new System.Drawing.Size(148, 17);
            this.chkOpenWhenDone.TabIndex = 13;
            this.chkOpenWhenDone.Text = "Open in Word when done";
            this.chkOpenWhenDone.UseVisualStyleBackColor = true;
            // 
            // btnValidateXml
            // 
            this.btnValidateXml.Location = new System.Drawing.Point(12, 320);
            this.btnValidateXml.Name = "btnValidateXml";
            this.btnValidateXml.Size = new System.Drawing.Size(183, 23);
            this.btnValidateXml.TabIndex = 14;
            this.btnValidateXml.Text = "Validate XML Files";
            this.btnValidateXml.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnValidateXml.UseVisualStyleBackColor = true;
            this.btnValidateXml.Click += new System.EventHandler(this.btnValidateXml_Click);
            // 
            // scroller
            // 
            this.scroller.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.scroller.FormattingEnabled = true;
            this.scroller.Location = new System.Drawing.Point(462, 123);
            this.scroller.MaxLines = 1000;
            this.scroller.Name = "scroller";
            this.scroller.Size = new System.Drawing.Size(392, 368);
            this.scroller.TabIndex = 7;
            this.scroller.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(863, 507);
            this.Controls.Add(this.btnValidateXml);
            this.Controls.Add(this.chkOpenWhenDone);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnGetCaveList);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmbBook);
            this.Controls.Add(this.scroller);
            this.Controls.Add(this.btnCreateDocx);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtFolder);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnScreenScrape);
            this.Controls.Add(this.tree);
            this.Controls.Add(this.btnLogin);
            this.Controls.Add(this.btnDownloadXml);
            this.MinimumSize = new System.Drawing.Size(879, 546);
            this.Name = "Form1";
            this.Text = "Northern Caves Data Cruncher";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.Button btnDownloadXml;
    private System.Windows.Forms.Button btnLogin;
    private System.Windows.Forms.TreeView tree;
    private System.Windows.Forms.Button btnScreenScrape;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.TextBox txtFolder;
    private System.Windows.Forms.Button btnBrowse;
    private System.Windows.Forms.Button btnCreateDocx;
    private Scroller scroller;
    private System.Windows.Forms.ComboBox cmbBook;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.Button btnGetCaveList;
    private System.Windows.Forms.Label label3;
    private System.Windows.Forms.CheckBox chkOpenWhenDone;
    private System.Windows.Forms.Button btnValidateXml;
  }
}

